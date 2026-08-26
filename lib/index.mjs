// dsh-office-com — COM 驱动真实 Office/WPS 实例的 DeepSeek Harness 原生插件
//
// 定位（区别于列表内所有文件级读写的 Office 插件）：
//   通过本机已装的 OfficeMCP（fastmcp + officemcp + pywin32）的 SSE 服务，
//   用 RunPython(code, data) 万能后门驱动真实 Excel/Word 实例：
//   VBA 宏 / 透视表 / 公式重算 / 已有文档深度排版。
//
// 三层架构：
//   底层  OfficeMCP SSE（本机 workbuddy Py3.13 已装）
//   中层  自举层：探测 officemcp → 起 SSE（动态端口 + 锁文件 singleton）→ 健康检查
//   上层  语义化工具层：ctx.tools.register 把 RunPython 封装成领域工具
//   兜底  检测失败静默降级纯 JS（exceljs/docx），不许炸插件树
//
// 环境变量：
//   OFFICE_PYTHON   覆盖 python 命令（默认探测 workbuddy Py3.13.12）
//   OFFICE_FOLDER   OfficeMCP 工作目录（默认 D:/OfficeMCP）

import { spawn, spawnSync } from 'node:child_process'
import { existsSync, readFileSync, writeFileSync, mkdirSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'
import net from 'node:net'

export const name = 'dsh-office-com'
export const inject = ['tools']

const HOST = '127.0.0.1'
const OFFICE_FOLDER = process.env.OFFICE_FOLDER || 'D:/OfficeMCP'
const STATE_DIR = join(homedir(), '.dsh-office-com')
const LOCK_FILE = join(STATE_DIR, 'office.lock')
const PROBE_TIMEOUT = 20000
const CALL_TIMEOUT = 90000
const HEALTH_TIMEOUT = 1500

const PROBE_CODE = 'from officemcp.OfficeMCP import RunOfficeMCP'

// ── 自举：干净 base python（三原则）──────────────────────────

/** 剥离 PYTHONPATH（Hermes/坏 venv 透传会劫持 import） */
function cleanEnv() {
  const env = { ...process.env }
  delete env.PYTHONPATH
  return env
}

function findPython() {
  const candidates = [
    process.env.OFFICE_PYTHON,
    join(homedir(), '.workbuddy', 'binaries', 'python', 'versions', '3.13.12', 'python.exe'),
    'python',
    'py',
  ].filter(Boolean)
  for (const c of candidates) {
    const args = c === 'py' ? ['-3', '-c', PROBE_CODE] : ['-c', PROBE_CODE]
    const r = spawnSync(c, args, { env: cleanEnv(), windowsHide: true, stdio: 'ignore' })
    if (r.status === 0) return c
  }
  return null
}

// ── SSE 服务单例：动态端口 + 锁文件 ─────────────────────────

function freePort() {
  return new Promise((resolve, reject) => {
    const srv = net.createServer()
    srv.on('error', reject)
    srv.listen(0, HOST, () => {
      const port = srv.address().port
      srv.close(() => resolve(port))
    })
  })
}

function tcpHealthy(port) {
  return new Promise((resolve) => {
    const sock = net.connect({ port, host: HOST })
    const timer = setTimeout(() => { sock.destroy(); resolve(false) }, HEALTH_TIMEOUT)
    sock.on('connect', () => { clearTimeout(timer); sock.destroy(); resolve(true) })
    sock.on('error', () => { clearTimeout(timer); resolve(false) })
  })
}

function readLock() {
  try { return JSON.parse(readFileSync(LOCK_FILE, 'utf-8')) } catch { return null }
}

function writeLock(rec) {
  try {
    mkdirSync(STATE_DIR, { recursive: true })
    writeFileSync(LOCK_FILE, JSON.stringify(rec))
  } catch { /* 锁写失败不阻塞 */ }
}

let serverChild = null // 本进程 spawn 的 SSE 子进程（退出时回收，Windows 下父进程退出不会自动杀子进程）

function killServer() {
  if (serverChild) {
    try { serverChild.kill() } catch { /* 已退出 */ }
    serverChild = null
  }
}
process.on('exit', killServer)

function spawnServer(py, port) {
  const script = [
    'import sys',
    `sys.argv=["officemcp","sse","--port","${port}","--folder","${OFFICE_FOLDER}"]`,
    'from officemcp.OfficeMCP import RunOfficeMCP',
    'RunOfficeMCP()',
  ].join(';')
  const child = spawn(py, ['-c', script], { env: cleanEnv(), windowsHide: true, stdio: ['ignore', 'ignore', 'pipe'] })
  serverChild = child
  let err = ''
  child.stderr.on('data', (d) => { err += d })
  child.on('error', () => {})
  child.on('exit', (code) => {
    if (serverChild === child) serverChild = null
    if (code && code !== 0) {
      // officemcp 收尾时可能带启动横幅退出（非零），那只是噪音；
      // 只有 stderr 像真错误（Traceback/Error/Exception…）才打全文
      const tail = err.trim().slice(-400)
      const looksError = /Traceback|Error|Exception|failed|Fatal|错误|异常/i.test(tail)
      if (looksError) console.error(`[dsh-office-com] officemcp exited ${code}: ${tail}`)
      else console.warn(`[dsh-office-com] officemcp exited ${code}（无错误文本，疑似进程收尾）`)
    }
  })
  return child.pid
}

async function waitHealthy(port) {
  const deadline = Date.now() + PROBE_TIMEOUT
  while (Date.now() < deadline) {
    if (await tcpHealthy(port)) return true
    await new Promise((r) => setTimeout(r, 200))
  }
  return false
}

let serverPromise = null // 单飞：并发 ensureServer 共享同一次 spawn（否则预热与首次调用并发会各起一个 SSE）

async function ensureServer() {
  if (!serverPromise) {
    serverPromise = spawnServerOnce().catch((e) => { serverPromise = null; throw e })
  }
  return serverPromise
}

async function spawnServerOnce() {
  const existing = readLock()
  if (existing && typeof existing.port === 'number' && (await tcpHealthy(existing.port))) {
    return existing.port
  }
  const py = findPython()
  if (!py) throw new Error('officemcp python not found (need OfficeMCP installed)')
  const port = await freePort()
  const pid = spawnServer(py, port)
  if (!(await waitHealthy(port))) throw new Error(`officemcp SSE failed to start on ${port}`)
  writeLock({ port, pid })
  console.log(`[dsh-office-com] officemcp SSE up on ${HOST}:${port} (pid ${pid})`)
  return port
}

// ── 最小 MCP-over-SSE 客户端（零依赖）────────────────────────

class OfficeClient {
  constructor(port) {
    this.base = `http://${HOST}:${port}`
    this.msgUrl = null
    this._id = 0
    this._pending = new Map()
    this._endpoint = new Promise((res, rej) => { this._endpointResolve = res; this._endpointReject = rej })
  }

  async connect(timeout = PROBE_TIMEOUT) {
    const res = await fetch(`${this.base}/sse`)
    if (!res.ok || !res.body) throw new Error(`SSE GET ${res.status}`)
    this._readLoop(res.body)
    await Promise.race([
      this._endpoint,
      new Promise((_, rej) => setTimeout(() => rej(new Error('SSE endpoint timeout')), timeout)),
    ])
    await this.call('initialize', {
      protocolVersion: '2024-11-05',
      capabilities: {},
      clientInfo: { name: 'dsh-office-com', version: '0.1.0' },
    })
    this._notify('notifications/initialized')
    return this
  }

  async _readLoop(body) {
    const reader = body.getReader()
    const dec = new TextDecoder()
    let buf = ''
    try {
      while (true) {
        const { done, value } = await reader.read()
        if (done) break
        buf += dec.decode(value, { stream: true })
        buf = buf.replace(/\r\n/g, '\n') // SSE 服务端常发 \r\n，统一成 \n 再切事件
        let i
        while ((i = buf.indexOf('\n\n')) >= 0) {
          this._onEvent(buf.slice(0, i))
          buf = buf.slice(i + 2)
        }
      }
    } catch { /* 流关闭 */ }
  }

  _onEvent(block) {
    let ev = 'message'
    let data = ''
    for (const line of block.split('\n')) {
      if (line.startsWith('event:')) ev = line.slice(6).trim()
      else if (line.startsWith('data:')) data += line.slice(5).trim()
    }
    if (ev === 'endpoint' && data) {
      this.msgUrl = data
      this._endpointResolve?.(data)
    } else if (ev === 'message' && data) {
      try {
        const m = JSON.parse(data)
        if (m.id != null && this._pending.has(m.id)) {
          const p = this._pending.get(m.id)
          this._pending.delete(m.id)
          m.error ? p.reject(new Error(JSON.stringify(m.error))) : p.resolve(m.result)
        }
      } catch { /* 忽略无法解析的事件 */ }
    }
  }

  async call(method, params) {
    if (!this.msgUrl) throw new Error('SSE not connected')
    const id = ++this._id
    const p = new Promise((resolve, reject) => {
      this._pending.set(id, { resolve, reject })
      setTimeout(() => { if (this._pending.delete(id)) reject(new Error(`MCP call timeout: ${method}`)) }, CALL_TIMEOUT)
    })
    await fetch(`${this.base}${this.msgUrl}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ jsonrpc: '2.0', id, method, params }),
    })
    return p
  }

  async _notify(method, params = {}) {
    try {
      await fetch(`${this.base}${this.msgUrl}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ jsonrpc: '2.0', method, params }),
      })
    } catch { /* notification 失败忽略 */ }
  }
}

// ── RunPython 桥 + 懒连接 ──────────────────────────────────

let client = null
let degraded = false
let degradeReason = ''

async function getClient() {
  if (client) return client
  const port = await ensureServer()
  client = await new OfficeClient(port).connect()
  return client
}

/** 调用 officemcp 的 RunPython 万能后门，返回 {success, output|error} */
async function runPython(code, data = '') {
  if (degraded) return { success: false, error: degradeReason }
  try {
    const c = await getClient()
    const result = await c.call('tools/call', { name: 'RunPython', arguments: { code, data } })
    const text = (result?.content || []).map((b) => b.text || '').join('\n')
    if (result?.isError) return { success: false, error: text || 'RunPython isError' }
    try {
      const parsed = JSON.parse(text)
      // officemcp 的 RunPython 返回 {success, output} 信封；本插件 CODE_* 常把 output 预编码成 JSON 字符串，
      // FastMCP 再序列化一层 → 这里解包一层，工具才能拿到结构化对象（否则 excel_new 之类返回的是 JSON 字符串）
      if (parsed && typeof parsed === 'object' && typeof parsed.output === 'string') {
        try { parsed.output = JSON.parse(parsed.output) } catch { /* 非 JSON 字符串（如 hello-from-COM）原样保留 */ }
      }
      return parsed
    } catch { return { success: false, error: text.slice(0, 500) || 'empty output' } }
  } catch (e) {
    return { success: false, error: e.message }
  }
}

/** 拼 COM 代码：args 以 JSON 走 data，代码内 json.loads(data) 取值。CoInitialize 是 COM Dispatch 的线程前置（否则 -2147221008）。
 *  path 统一归一化为反斜杠——实测 Excel SaveAs/Open 对正斜杠路径解析异常（会把 D:/x 拼成乱路径）。 */
function comCode(body) {
  return ['import json, pythoncom', 'pythoncom.CoInitialize()', 'args = json.loads(data) if data else {}', "if args.get('path'): args['path'] = args['path'].replace('/', '\\\\')", ...body].join('\n')
}

// ── 工具：render/execute ─────────────────────────────────────

function renderText(_args, value) {
  if (!value) return [{ type: 'text', text: 'dsh-office-com 失败' }]
  if (value.ok === false || value.success === false) return [{ type: 'text', text: `失败: ${value.error}` }]
  return [{ type: 'text', text: typeof value.output === 'string' ? value.output : JSON.stringify(value.output ?? value) }]
}

async function executeRunCode(code, data) {
  const r = await runPython(code, data)
  if (r.success) return { ok: true, output: r.output }
  return { ok: false, error: r.error }
}

function reg(ctx, tool) {
  // DSH 要求 parameters 是完整 JSON Schema（{type:'object', properties, required}），
  // 不能裸传属性表（否则校验报 schema must be type:'object', got 'type: null'）。
  const props = tool.parameters || {}
  const required = Object.keys(props).filter((k) => props[k]?.required === true)
  for (const k of Object.keys(props)) delete props[k].required // required 归顶层数组，属性内不留非标准字段
  ctx.tools.register({
    ...tool,
    parameters: { type: 'object', properties: props, required },
    output: { schema: { type: 'object' }, render: renderText },
  })
}

// office_apps — 发现 Office 应用
const CODE_OFFICE_APPS = comCode([
  "apps = {'excel': Officer.IsAppAvailable('Excel'), 'word': Officer.IsAppAvailable('Word'), 'outlook': Officer.IsAppAvailable('Outlook'), 'running': Officer.RunningApps()}",
  "output = json.dumps(apps, ensure_ascii=False)",
])

// excel_formula_set — 写公式（COM Range.Formula，原样写入）
const CODE_FORMULA_SET = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  "if args.get('path'):",
  "    wb = excel.Workbooks.Open(args['path'])",
  'else:',
  '    wb = excel.ActiveWorkbook',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  "ws = wb.Worksheets(args['sheet']) if args.get('sheet') else wb.ActiveSheet",
  "ws.Range(args['range']).Formula = args['formula']",
  "output = json.dumps({'formula_set': '%s!%s = %s' % (ws.Name, args['range'], args['formula'])}, ensure_ascii=False)",
])

// excel_recalc — 强制重算 + 取计算后活值（日期/NaN 清洗同 read_range）
const CODE_RECALC = comCode([
  'import datetime, math',
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  "if args.get('path'):",
  "    wb = excel.Workbooks.Open(args['path'])",
  'else:',
  '    wb = excel.ActiveWorkbook',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  'excel.Calculate()',
  'def san(v):',
  '    if v is None: return None',
  '    if isinstance(v, (datetime.datetime, datetime.date)): return v.isoformat()',
  '    if isinstance(v, float) and not math.isfinite(v): return None',
  '    return v',
  'def tolist(x):',
  '    if isinstance(x, (tuple, list)): return [tolist(i) for i in x]',
  '    return san(x)',
  "if args.get('range'):",
  "    ws = wb.Worksheets(args['sheet']) if args.get('sheet') else wb.ActiveSheet",
  "    v = tolist(ws.Range(args['range']).Value)",
  "    output = json.dumps({'recalc': 'ok', 'range': args['range'], 'value': v}, ensure_ascii=False)",
  'else:',
  "    output = json.dumps({'recalc': 'ok'}, ensure_ascii=False)",
])

// excel_vba_run — 运行已有宏 / 注入执行
const CODE_VBA_RUN = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  "if args.get('path'):",
  "    wb = excel.Workbooks.Open(args['path'])",
  'else:',
  '    wb = excel.ActiveWorkbook',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  "r = excel.Run(args['macro'])",
  "output = json.dumps({'vba_run': args['macro'], 'result': r}, ensure_ascii=False, default=str)",
])

// office_launch — 显式启动应用（镜像 officemcp Launch 实现：Application + Visible）
const CODE_LAUNCH = comCode([
  "app_name = args.get('app', 'Excel')",
  'app = Officer.Application(app_name)',
  "if app is None: raise Exception('应用不可用: ' + app_name)",
  "app.Visible = bool(args.get('visible', True))",
  "output = json.dumps({'launched': True, 'app': app_name}, ensure_ascii=False)",
])

// excel_new — 新建空工作簿
const CODE_NEW = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  'wb = excel.Workbooks.Add()',
  "sheets = [ws.Name for ws in wb.Worksheets]",
  "output = json.dumps({'new_workbook': True, 'sheets': sheets}, ensure_ascii=False)",
])

// excel_open — 打开已有工作簿，返回结构信息
const CODE_OPEN = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  "if not args.get('path'): raise Exception('缺少 path')",
  'wb = excel.Workbooks.Open(args["path"])',
  "sheets = [ws.Name for ws in wb.Worksheets]",
  "output = json.dumps({'path': args['path'], 'active_sheet': wb.ActiveSheet.Name, 'sheets': sheets}, ensure_ascii=False)",
])

// excel_read_range — 读区域活值（2D 数组；日期→iso，NaN/Inf→null）
const CODE_READ_RANGE = comCode([
  'import datetime, math',
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  "if args.get('path'):",
  "    wb = excel.Workbooks.Open(args['path'])",
  'else:',
  '    wb = excel.ActiveWorkbook',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  "ws = wb.Worksheets(args['sheet']) if args.get('sheet') else wb.ActiveSheet",
  'def san(v):',
  '    if v is None: return None',
  '    if isinstance(v, (datetime.datetime, datetime.date)): return v.isoformat()',
  '    if isinstance(v, float) and not math.isfinite(v): return None',
  '    return v',
  'def tolist(x):',
  '    if isinstance(x, (tuple, list)): return [tolist(i) for i in x]',
  '    return san(x)',
  "output = json.dumps({'range': args['range'], 'value': tolist(ws.Range(args['range']).Value)}, ensure_ascii=False)",
])

// excel_write_range — 写值（标量或 2D 数组；list→tuple 供 COM 赋值）
const CODE_WRITE_RANGE = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  "if args.get('path'):",
  "    wb = excel.Workbooks.Open(args['path'])",
  'else:',
  '    wb = excel.ActiveWorkbook',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  "ws = wb.Worksheets(args['sheet']) if args.get('sheet') else wb.ActiveSheet",
  'def tocom(v):',
  '    if isinstance(v, list): return tuple(tocom(i) for i in v)',
  '    return v',
  'val = tocom(args.get("value"))',
  'ws.Range(args["range"]).Value = val',
  "output = json.dumps({'range': args['range'], 'written': True}, ensure_ascii=False)",
])

// word_open — 打开已有 Word 文档，返回结构信息（段落数 + 文本预览）
const CODE_WORD_OPEN = comCode([
  'word = Officer.Word',
  'word.DisplayAlerts = 0',
  "if not word: raise Exception('Word 未安装/不可用')",
  "if not args.get('path'): raise Exception('缺少 path')",
  'doc = word.Documents.Open(args["path"])',
  'text = doc.Content.Text',
  "output = json.dumps({'path': args['path'], 'paragraphs': doc.Paragraphs.Count, 'chars': len(text), 'preview': text[:300]}, ensure_ascii=False)",
  'doc.Close(False)',
])

// word_edit — 全文查找替换（纯文本 Content.Text 替换，走真实 Word 实例读写，保存回原文件）
const CODE_WORD_EDIT = comCode([
  'word = Officer.Word',
  'word.DisplayAlerts = 0',
  "if not word: raise Exception('Word 未安装/不可用')",
  "if args.get('path'):",
  "    doc = word.Documents.Open(args['path'])",
  'else:',
  '    doc = word.ActiveDocument',
  "if doc is None: raise Exception('无活动文档且未给 path')",
  'find_text = args["find"]',
  "replace_text = args.get('replace', '')",
  'text = doc.Content.Text',
  'count = text.count(find_text)',
  'if count > 0:',
  '    doc.Content.Text = text.replace(find_text, replace_text)',
  '    doc.Save()',
  'doc.Close(False)',
  "output = json.dumps({'find': find_text, 'replace': replace_text, 'count': count, 'saved': count > 0}, ensure_ascii=False)",
])

// excel_pivot_create — 创建透视表（PivotCache + PivotTable，行/列/值字段）
const CODE_PIVOT_CREATE = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  "if args.get('path'):",
  "    wb = excel.Workbooks.Open(args['path'])",
  'else:',
  '    wb = excel.ActiveWorkbook',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  "if not args.get('range'): raise Exception('缺少 range（源数据区域，含表头）')",
  "ws = wb.Worksheets(args['sheet']) if args.get('sheet') else wb.ActiveSheet",
  'ws_dst = wb.Worksheets.Add()',
  'n = 0',
  'for w in wb.Worksheets:',
  '    n += len(list(w.PivotTables()))',
  "tname = 'PivotTable%d' % (n + 1)",
  "pc = wb.PivotCaches().Create(SourceType=1, SourceData=ws.Range(args['range']))", // SourceType=1 xlDatabase
  "pt = pc.CreatePivotTable(TableDestination=ws_dst.Range('A1'), TableName=tname)",
  "for f in (args.get('rows') or []):",
  '    pt.PivotFields(f).Orientation = 1', // xlRowField
  "for f in (args.get('columns') or []):",
  '    pt.PivotFields(f).Orientation = 2', // xlColumnField
  "for f in (args.get('values') or []):",
  "    pt.AddDataField(pt.PivotFields(f), f + '_求和', -4157)", // xlSum
  "output = json.dumps({'pivot': 'ok', 'dst_sheet': ws_dst.Name, 'rows': args.get('rows'), 'columns': args.get('columns'), 'values': args.get('values')}, ensure_ascii=False)",
])

// excel_pivot_refresh — 刷新工作簿内所有透视表
const CODE_PIVOT_REFRESH = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  "if args.get('path'):",
  "    wb = excel.Workbooks.Open(args['path'])",
  'else:',
  '    wb = excel.ActiveWorkbook',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  'count = 0',
  'for ws in wb.Worksheets:',
  '    for pt in ws.PivotTables():',
  '        pt.RefreshTable()',
  '        count += 1',
  "output = json.dumps({'refreshed': count}, ensure_ascii=False)",
])

// excel_journal_post — 写会计分录到账簿 + 借贷平衡校验（错则标红）
const CODE_JOURNAL_POST = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  "if args.get('path'):",
  "    wb = excel.Workbooks.Open(args['path'])",
  'else:',
  '    wb = excel.ActiveWorkbook',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  "ws = wb.Worksheets(args['sheet']) if args.get('sheet') else wb.ActiveSheet",
  'last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row', // xlUp 找 A 列最后一行
  "if last == 1 and str(ws.Cells(1,1).Value or '').strip() == '':",
  "    ws.Cells(1,1).Value = '日期'",
  "    ws.Cells(1,2).Value = '摘要'",
  "    ws.Cells(1,3).Value = '科目'",
  "    ws.Cells(1,4).Value = '借方'",
  "    ws.Cells(1,5).Value = '贷方'",
  '    start = 2',
  'else:',
  '    start = last + 1',
  'dr = 0.0; cr = 0.0',
  "entries = args.get('entries') or []",
  'for i, e in enumerate(entries):',
  '    r = start + i',
  "    ws.Cells(r,1).Value = e.get('date','')",
  "    ws.Cells(r,2).Value = e.get('desc','')",
  "    ws.Cells(r,3).Value = e.get('account','')",
  '    d = float(e.get("debit") or 0); c = float(e.get("credit") or 0)',
  '    ws.Cells(r,4).Value = d',
  '    ws.Cells(r,5).Value = c',
  '    dr += d; cr += c',
  'balanced = abs(dr - cr) < 1e-9',
  'if not balanced and entries:',
  '    ws.Range(ws.Cells(start,1), ws.Cells(start+len(entries)-1,5)).Font.Color = 255', // 红色标错
  "output = json.dumps({'posted': len(entries), 'debit_total': dr, 'credit_total': cr, 'balanced': balanced}, ensure_ascii=False)",
])

// excel_ledger_gen — 日记账 → 科目总账（按科目聚合借/贷，余额=借-贷公式）
const CODE_LEDGER_GEN = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  "if args.get('path'):",
  "    wb = excel.Workbooks.Open(args['path'])",
  'else:',
  '    wb = excel.ActiveWorkbook',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  "ws = wb.Worksheets(args['journal_sheet']) if args.get('journal_sheet') else wb.ActiveSheet",
  'last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row',
  "if last < 2: raise Exception('日记账为空（无数据行）')",
  'accounts = {}',
  'for r in range(2, last+1):',
  '    acc = ws.Cells(r,3).Value',
  '    if acc is None: continue',
  '    d = float(ws.Cells(r,4).Value or 0); c = float(ws.Cells(r,5).Value or 0)',
  '    if acc not in accounts: accounts[acc] = [0.0, 0.0]',
  '    accounts[acc][0] += d; accounts[acc][1] += c',
  "ws_out = wb.Worksheets(args['output_sheet']) if args.get('output_sheet') else wb.Worksheets.Add()",
  "ws_out.Cells(1,1).Value = '科目'; ws_out.Cells(1,2).Value = '借方合计'; ws_out.Cells(1,3).Value = '贷方合计'; ws_out.Cells(1,4).Value = '余额'",
  'r = 2',
  'for acc in sorted(accounts.keys()):',
  '    d, c = accounts[acc]',
  '    ws_out.Cells(r,1).Value = acc',
  '    ws_out.Cells(r,2).Value = d',
  '    ws_out.Cells(r,3).Value = c',
  "    ws_out.Cells(r,4).Formula = '=B%d-C%d' % (r, r)",
  '    r += 1',
  "output = json.dumps({'accounts': len(accounts), 'output_sheet': ws_out.Name}, ensure_ascii=False)",
])

// ── 插件入口 ─────────────────────────────────────────────────

// 导出内部函数供冒烟测试（DSH 只用 name/inject/apply）
export { findPython, ensureServer, runPython }

/** @param {import('@deepseek-ai/cordis').Context} ctx */
export function apply(ctx) {
  // 启动即探：officemcp 可用性决定是否降级（不阻塞 apply）
  const py = findPython()
  if (!py) {
    degraded = true
    degradeReason = '本机未装 OfficeMCP（officemcp），已降级为纯 JS（暂只读，写/COM 工具不可用）'
    console.warn(`[dsh-office-com] ${degradeReason}`)
  } else {
    // 后台预热 SSE（失败不致命，工具调用时还会重试一次）
    ensureServer().catch((e) => {
      console.warn(`[dsh-office-com] 预热失败（懒启动兜底）: ${e.message}`)
    })
  }

  reg(ctx, {
    name: 'office_apps',
    description: '发现本机可用的 Microsoft Office 应用（Excel/Word/Outlook）与正在运行的实例。',
    parameters: {},
    execute: () => executeRunCode(CODE_OFFICE_APPS, ''),
  })

  reg(ctx, {
    name: 'excel_formula_set',
    description: '在真实 Excel 实例中向单元格/区域写入公式（原样写入 =SUM() 等，交给 Excel 引擎计算）。path 给文件路径则先打开，否则用活动工作簿。',
    parameters: {
      path: { type: 'string', required: false, description: '工作簿完整路径（可选，缺省用活动工作簿）' },
      sheet: { type: 'string', required: false, description: '工作表名（可选，缺省用活动表）' },
      range: { type: 'string', required: true, description: '目标区域，如 A1 或 A1:B10' },
      formula: { type: 'string', required: true, description: '公式内容，如 =SUM(A1:A10)' },
    },
    execute: (args) => executeRunCode(CODE_FORMULA_SET, JSON.stringify(args)),
  })

  reg(ctx, {
    name: 'excel_recalc',
    description: '强制 Excel 重算全部公式并取计算后的活值（COM 是活公式链，区别于 openpyxl 写死数值的死穴）。',
    parameters: {
      path: { type: 'string', required: false, description: '工作簿完整路径（可选）' },
      sheet: { type: 'string', required: false, description: '工作表名（可选）' },
      range: { type: 'string', required: false, description: '取值的区域，如 A1（可选，缺省只重算不取值）' },
    },
    execute: (args) => executeRunCode(CODE_RECALC, JSON.stringify(args)),
  })

  reg(ctx, {
    name: 'excel_vba_run',
    description: '运行工作簿中已有的 VBA 宏（Application.Run），或在已打开实例上执行指定宏。',
    parameters: {
      path: { type: 'string', required: false, description: '工作簿完整路径（可选，缺省用活动工作簿）' },
      macro: { type: 'string', required: true, description: '宏名，如 Module1.MyMacro' },
    },
    execute: (args) => executeRunCode(CODE_VBA_RUN, JSON.stringify(args)),
  })

  reg(ctx, {
    name: 'office_launch',
    description: '显式启动本机的 Microsoft Office 应用（Excel/Word/Outlook）。',
    parameters: {
      app: { type: 'string', required: false, description: '应用名：Excel / Word / Outlook（缺省 Excel）' },
      visible: { type: 'boolean', required: false, description: '窗口是否可见（缺省 true）' },
    },
    execute: (args) => executeRunCode(CODE_LAUNCH, JSON.stringify(args)),
  })

  reg(ctx, {
    name: 'excel_new',
    description: '新建一个空 Excel 工作簿，返回工作表清单（作为后续写值/公式的目标）。',
    parameters: {},
    execute: () => executeRunCode(CODE_NEW, ''),
  })

  reg(ctx, {
    name: 'excel_open',
    description: '打开已有 Excel 工作簿，返回工作表清单与活动表名。',
    parameters: {
      path: { type: 'string', required: true, description: '工作簿完整路径，如 C:/path/book.xlsx' },
    },
    execute: (args) => executeRunCode(CODE_OPEN, JSON.stringify(args)),
  })

  reg(ctx, {
    name: 'excel_read_range',
    description: '读取 Excel 区域当前值（公式返回计算后活值），返回 2D 数组。',
    parameters: {
      path: { type: 'string', required: false, description: '工作簿完整路径（可选，缺省活动工作簿）' },
      sheet: { type: 'string', required: false, description: '工作表名（可选，缺省活动表）' },
      range: { type: 'string', required: true, description: '目标区域，如 A1 或 A1:C10' },
    },
    execute: (args) => executeRunCode(CODE_READ_RANGE, JSON.stringify(args)),
  })

  reg(ctx, {
    name: 'excel_write_range',
    description: '向 Excel 区域写入值：标量或与区域匹配的 2D 数组（公式请用 excel_formula_set）。',
    parameters: {
      path: { type: 'string', required: false, description: '工作簿完整路径（可选）' },
      sheet: { type: 'string', required: false, description: '工作表名（可选）' },
      range: { type: 'string', required: true, description: '目标区域，如 A1 或 A1:C3' },
      value: { required: true, description: '写入的值：标量，或与区域形状一致的 2D 数组' },
    },
    execute: (args) => executeRunCode(CODE_WRITE_RANGE, JSON.stringify(args)),
  })

  reg(ctx, {
    name: 'word_open',
    description: '打开已有 Word 文档，返回段落数、字符数与开头文本预览（COM 读取真实排版结构）。',
    parameters: {
      path: { type: 'string', required: true, description: 'Word 文档完整路径，如 C:/path/doc.docx' },
    },
    execute: (args) => executeRunCode(CODE_WORD_OPEN, JSON.stringify(args)),
  })

  reg(ctx, {
    name: 'word_edit',
    description: '在 Word 文档中全文查找替换（COM Find，wdReplaceAll）。path 给文件则先打开，否则用活动文档。',
    parameters: {
      path: { type: 'string', required: false, description: 'Word 文档完整路径（可选，缺省活动文档）' },
      find: { type: 'string', required: true, description: '要查找的文本' },
      replace: { type: 'string', required: false, description: '替换成的文本（缺省删除）' },
    },
    execute: (args) => executeRunCode(CODE_WORD_EDIT, JSON.stringify(args)),
  })

  reg(ctx, {
    name: 'excel_pivot_create',
    description: '基于源数据区域创建 Excel 透视表（真实 PivotCache/PivotTable），返回新建透视表所在工作表。值字段按求和汇总。',
    parameters: {
      path: { type: 'string', required: false, description: '工作簿完整路径（可选，缺省活动工作簿）' },
      sheet: { type: 'string', required: false, description: '源数据工作表名（可选，缺省活动表）' },
      range: { type: 'string', required: true, description: '源数据区域（含表头），如 A1:C100' },
      rows: { type: 'array', items: { type: 'string' }, required: true, description: '行字段名数组，如 ["科目"]' },
      columns: { type: 'array', items: { type: 'string' }, required: false, description: '列字段名数组，如 ["月份"]' },
      values: { type: 'array', items: { type: 'string' }, required: true, description: '值字段名数组（求和），如 ["金额"]' },
    },
    execute: (args) => executeRunCode(CODE_PIVOT_CREATE, JSON.stringify(args)),
  })

  reg(ctx, {
    name: 'excel_pivot_refresh',
    description: '刷新工作簿内所有透视表（源数据变化后取最新汇总）。',
    parameters: {
      path: { type: 'string', required: false, description: '工作簿完整路径（可选，缺省活动工作簿）' },
    },
    execute: (args) => executeRunCode(CODE_PIVOT_REFRESH, JSON.stringify(args)),
  })

  reg(ctx, {
    name: 'excel_journal_post',
    description: '把会计分录写入 Excel 账簿（日期/摘要/科目/借方/贷方，缺表头自动建），并做借贷平衡校验（借≠贷标红）。会计旗舰场景第一环。',
    parameters: {
      path: { type: 'string', required: false, description: '工作簿完整路径（可选，缺省活动工作簿）' },
      sheet: { type: 'string', required: false, description: '账簿工作表名（可选，缺省活动表）' },
      entries: {
        type: 'array',
        required: true,
        description: '分录数组，每项 {date, desc, account, debit, credit}',
        items: {
          type: 'object',
          properties: {
            date: { type: 'string', description: '日期' },
            desc: { type: 'string', description: '摘要' },
            account: { type: 'string', description: '科目名称' },
            debit: { type: 'number', description: '借方金额' },
            credit: { type: 'number', description: '贷方金额' },
          },
        },
      },
    },
    execute: (args) => executeRunCode(CODE_JOURNAL_POST, JSON.stringify(args)),
  })

  reg(ctx, {
    name: 'excel_ledger_gen',
    description: '从日记账生成科目总账：按科目聚合借贷（Python 端 SUMIF 等价），输出 科目/借方合计/贷方合计/余额（公式）。',
    parameters: {
      path: { type: 'string', required: false, description: '工作簿完整路径（可选）' },
      journal_sheet: { type: 'string', required: false, description: '日记账工作表名（可选，缺省活动表）' },
      output_sheet: { type: 'string', required: false, description: '总账输出工作表名（可选，缺省新建）' },
    },
    execute: (args) => executeRunCode(CODE_LEDGER_GEN, JSON.stringify(args)),
  })

  console.log(`[dsh-office-com] plugin loaded${degraded ? ' (degraded)' : ''}`)
}
