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
let ownsServer = false // 本进程是不是这个 SSE 服务的发起者（附着别人的服务就无权收尾）

function killServer() {
  if (serverChild) {
    try { serverChild.kill() } catch { /* 已退出 */ }
    serverChild = null
  }
}

// ── 退出收尾：回收我们起的不可见 Office 实例 ─────────────────
// officemcp 的 Officer.Excel 是 GetActiveObject-else-Dispatch 出来的实例，被永久缓存在 python 进程上；
// 我们退出时只 kill python SSE 子进程（Windows 上是 TerminateProcess，atexit 不跑），而 Excel 只要
// 还开着工作簿就不会自己退 → 宿主退出后留下**不可见**的孤儿 EXCEL.EXE（占内存 + 占文件锁）。
// 收尾判据：启动前不存在（不是用户开的）+ 当前不可见（没人正在看）→ 才 Quit。
// 用户自己开的实例、或被切到可见的实例一律不动。
const APP_IMAGES = [['EXCEL.EXE', 'Excel'], ['WINWORD.EXE', 'Word'], ['POWERPNT.EXE', 'PowerPoint']]

function pidsOfImage(image) {
  const r = spawnSync('tasklist', ['/FI', `IMAGENAME eq ${image}`, '/FO', 'CSV', '/NH'], { encoding: 'utf8' })
  return (r.stdout || '').split('\n').map((l) => l.match(/^"[^"]+","(\d+)"/)).filter(Boolean).map((m) => Number(m[1]))
}

function runningOfficeImages() {
  return APP_IMAGES.map(([img]) => img).filter((img) => pidsOfImage(img).length > 0)
}

function cleanupOwnedApps() {
  if (!ownsServer) return // 附着别人的 SSE：那些实例不是我们起的
  const rec = readLock()
  if (!rec || !Array.isArray(rec.preexisting)) return // 老锁文件没记基线，宁可不收尾也不误杀
  const owned = APP_IMAGES
    .filter(([img]) => !rec.preexisting.includes(img) && pidsOfImage(img).length > 0)
    .map(([img, app]) => `${app}|${img}`)
  if (!owned.length) return
  const py = findPython()
  if (!py) return
  // 先礼后兵：COM 客户端被 TerminateProcess 猝死后，Excel 会变成「留着 Hwnd、Visible=False、
  // 工作簿为 0，但 Quit() 也送不走」的僵尸，只能按 Hwnd 拿 pid 强制结束。
  // 强杀只对「不可见 + 启动前不存在」的实例发生——用户看得见的、或用户自己开的一律不动；
  // 且落地前复核 pid 背后还是同一个 exe（1.5s 内 pid 被复用的话，宁可放过也不误杀）。
  const script = [
    'import sys, time, pythoncom, win32com.client, win32process, win32api, win32con',
    'pythoncom.CoInitialize()',
    'for spec in sys.argv[1:]:',
    '    name, image = spec.split("|")',
    '    try:',
    '        app = win32com.client.GetActiveObject(name + ".Application")',
    '    except Exception:',
    '        continue',
    '    pid = None',
    '    try:',
    '        if app.Visible:',
    '            continue',
    '        pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]',
    '        app.DisplayAlerts = False',
    '        app.Quit()',
    '        time.sleep(1.5)',
    '    except Exception:',
    '        pass',
    '    if not pid:',
    '        continue',
    '    try:',
    '        h = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_TERMINATE, 0, pid)',
    '        if not win32process.GetModuleFileNameEx(h, 0).upper().endswith(image):',
    '            win32api.CloseHandle(h)',
    '            continue',
    '        win32api.TerminateProcess(h, 0)',
    '        win32api.CloseHandle(h)',
    '    except Exception:',
    '        pass',
  ].join('\n')
  try {
    spawnSync(py, ['-c', script, ...owned], { env: cleanEnv(), windowsHide: true, stdio: 'ignore', timeout: 10000 })
  } catch { /* 收尾失败不阻塞退出 */ }
}

// 收尾挂在 exit 上（正常退出与 Ctrl+C 默认终止都会走到它）。
// 不接 SIGINT/SIGTERM：插件活在宿主进程里，接信号会盖掉 DSH 自己的退出语义。
process.on('exit', () => {
  cleanupOwnedApps()
  killServer()
})

function spawnServer(py, port) {
  const script = [
    'import sys',
    `sys.argv=["officemcp","sse","--port","${port}","--folder","${OFFICE_FOLDER}"]`,
    'from officemcp.OfficeMCP import RunOfficeMCP',
    'RunOfficeMCP()',
  ].join(';')
  const child = spawn(py, ['-c', script], { env: cleanEnv(), windowsHide: true, stdio: ['ignore', 'ignore', 'pipe'] })
  serverChild = child
  ownsServer = true
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
  const preexisting = runningOfficeImages() // 动手前用户已经开着的 Office —— 收尾时必须放过它们
  const port = await freePort()
  const pid = spawnServer(py, port)
  if (!(await waitHealthy(port))) throw new Error(`officemcp SSE failed to start on ${port}`)
  writeLock({ port, pid, preexisting })
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
      clientInfo: { name: 'dsh-office-com', version: '0.1.1' },
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

/** SSE/COM 通道断了之后重置桥接单例：client 与 serverPromise 都作废，下次调用自动重连（而不是永远等超时） */
function resetBridge() {
  client = null
  serverPromise = null
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
    // 连接类错误（SSE 断/服务死/超时）→ 桥接失效，重置后下次调用自动重连
    if (/SSE not connected|MCP call timeout|fetch failed|ECONNREFUSED|network|socket hang up/i.test(e.message)) {
      resetBridge()
    }
    return { success: false, error: e.message }
  }
}

/** 拼 COM 代码：args 以 JSON 走 data，代码内 json.loads(data) 取值。CoInitialize 是 COM Dispatch 的线程前置（否则 -2147221008）。
 *  path 统一归一化为反斜杠——实测 Excel SaveAs/Open 对正斜杠路径解析异常（会把 D:/x 拼成乱路径）。 */
function comCode(body) {
  return [
    'import json, pythoncom',
    'pythoncom.CoInitialize()',
    'args = json.loads(data) if data else {}',
    "if args.get('path'): args['path'] = args['path'].replace('/', '\\\\')",
    '_existing_open = [False]',
    // 已在用户手里打开的文件：Excel/Word 的 Open 对同一个文件返回**同一个对象**，若照旧当成"自己开的"
    // 在 finally 里 Close(False)，就是关掉用户正在编辑的窗口并丢弃其未保存修改。先按 FullName 找一遍，
    // 命中则一律按附着处理（既不给关、也不给它存）——对应工单 §四「重点保护：当前用户正在编辑的工作簿」。
    'def _find_open(coll, path):',
    '    try:',
    "        want = path.replace('/', '\\\\').lower()",
    '        for d in coll:',
    '            try:',
    '                if str(d.FullName).lower() == want: return d',
    '            except Exception: pass',
    '    except Exception: pass',
    '    return None',
    // 工作簿解析助手：给 path 就 Open（返回 (wb, True)），否则用活动工作簿（(wb, False)）。
    // 只有自己 Open 的才在 finally 里 Close——否则带 path 反复调用会堆积打开的工作簿（文件锁/内存泄漏）。
    // 写操作（写值/公式/分录/透视/宏）在关闭前 Save()，读操作不保存，避免误改源文件。
    'def _resolve_wb(excel):',
    "    if args.get('path'):",
    '        hit = _find_open(excel.Workbooks, args["path"])',
    '        if hit is not None:',
    '            _existing_open[0] = True',
    '            return hit, False',
    "        return excel.Workbooks.Open(args['path']), True",
    '    return excel.ActiveWorkbook, False',
    'def _resolve_doc(word):',
    "    if args.get('path'):",
    '        hit = _find_open(word.Documents, args["path"])',
    '        if hit is not None:',
    '            _existing_open[0] = True',
    '            return hit, False',
    "        return word.Documents.Open(args['path']), True",
    '    return word.ActiveDocument, False',
    ...body,
    // body 之后收尾：碰的是用户自己在用的工作簿时，信封必须说实话——改进了内存但没落盘。
    // （finalize 现在只看 args.path 判 saved，这里把观测到的事实回传给上层。）
    'if _existing_open[0]:',
    '    try:',
    '        _o = json.loads(output)',
    '        if isinstance(_o, dict):',
    "            _o['attached_existing_open'] = True",
    '            output = json.dumps(_o, ensure_ascii=False)',
    '    except Exception: pass',
  ].join('\n')
}

// ── 统一结果协议 v0.2 ────────────────────────────────────────
// 所有工具返回同一信封，Agent 只需读 ok / error_code / retryable 就能分支，
// 不必解析每个工具各自不同的返回形状。
//   成功：{ ok, operation, changed, saved, verified, warnings, output }
//   失败：{ ok, operation, error_code, error, retryable, partial_changes }
// needsPath: 只有带 path（managed 模式，自己 Open/Close）才写盘；附着模式只改内存里的活动工作簿。

const TOOL_META = {
  office_apps:         { changed: false, saved: false, verified: false },
  office_launch:       { changed: false, saved: false, verified: false },
  excel_new:           { changed: false, saved: false, verified: false },
  excel_open:          { changed: false, saved: false, verified: false },
  excel_read_range:    { changed: false, saved: false, verified: false },
  excel_recalc:        { changed: false, saved: false, verified: true  }, // 重算后读回活值 = 已验证
  excel_formula_set:   { changed: true,  saved: true,  verified: false, needsPath: true },
  excel_write_range:   { changed: true,  saved: true,  verified: false, needsPath: true },
  excel_vba_run:       { changed: true,  saved: true,  verified: false, needsPath: true },
  excel_pivot_create:  { changed: true,  saved: true,  verified: false, needsPath: true },
  excel_pivot_refresh: { changed: true,  saved: true,  verified: false, needsPath: true },
  excel_journal_post:  { changed: true,  saved: true,  verified: true,  needsPath: true }, // 借贷平衡校验 = 已验证
  excel_ledger_gen:    { changed: true,  saved: true,  verified: false, needsPath: true },
  word_open:           { changed: false, saved: false, verified: false },
  word_edit:           { changed: true,  saved: true,  verified: false },
}

/** 把 Python 抛出的错误文本映射到稳定 error_code + 可重试性。 */
function classifyError(error) {
  const e = String(error || '')
  if (/SSE not connected|MCP call timeout|fetch failed|ECONNREFUSED|socket hang up|Failed to fetch|network/i.test(e)) {
    return { code: 'CHANNEL_UNAVAILABLE', retryable: true } // 通道断了，重连可重试
  }
  if (/降级|degraded|未装 OfficeMCP|not installed/i.test(e)) {
    return { code: 'DEGRADED', retryable: false }
  }
  if (/被占用|locked|sharing violation|permission denied|正在使用|in use/i.test(e)) {
    return { code: 'FILE_LOCKED', retryable: true } // 文件被占，等释放后重试
  }
  if (/源数据区域缺少|缺少字段|源数据/i.test(e)) {
    return { code: 'SOURCE_RANGE_INVALID', retryable: false }
  }
  if (/缺少 path|缺少 range|缺少 find|缺少 macro/i.test(e)) {
    return { code: 'MISSING_PARAM', retryable: false }
  }
  if (/无活动工作簿|无活动文档|ActiveWorkbook|ActiveDocument/i.test(e)) {
    return { code: 'NO_ACTIVE_DOCUMENT', retryable: false }
  }
  if (/未安装|不可用|not available|IsAppAvailable/i.test(e)) {
    return { code: 'APP_UNAVAILABLE', retryable: false }
  }
  if (/日记账为空|为空|empty/i.test(e)) {
    return { code: 'EMPTY_SOURCE', retryable: false }
  }
  if (/无法打开|cannot open|Open failed/i.test(e)) {
    return { code: 'OPEN_FAILED', retryable: false }
  }
  return { code: 'UNKNOWN', retryable: false }
}

/** 原始 {ok, output} / {ok:false, error} → 统一信封。 */
function finalize(operation, meta, args, r) {
  if (r && r.ok === true) {
    const out = r.output
    const warnings = []
    let verified = meta.verified ?? false
    // 会计特例：借贷不平衡 = 成功但带 warning（verified 置 false，不是错误是提示）
    if (out && typeof out === 'object' && out.balanced === false) {
      verified = false
      warnings.push('借贷不平衡：借方合计 ≠ 贷方合计')
    }
    return {
      ok: true,
      operation,
      changed: meta.changed ?? false,
      // needsPath 的工具：只有给了 path 才会 Save()；附着模式（无 path）改的是用户在用的工作簿，
      // 写进去了但**没落盘**——信封必须说实话，否则 Agent 会以为已经保存。
      saved: (meta.saved ?? false) && !(meta.needsPath && args?.path == null),
      verified,
      warnings,
      output: out,
    }
  }
  const err = (r && r.error) || 'unknown error'
  const cls = classifyError(err)
  return {
    ok: false,
    operation,
    error_code: cls.code,
    error: err,
    retryable: cls.retryable,
    // 写工具在 attached 模式（无 path）失败时，部分修改可能残留（无人 close 丢弃）；
    // managed 模式（有 path）finally 里 Close(False) 会丢弃，故 partial=false。
    partial_changes: !!(meta.changed && args?.path == null),
  }
}

function renderText(_args, value) {
  if (!value) return [{ type: 'text', text: 'dsh-office-com 失败' }]
  if (value.ok === false) return [{ type: 'text', text: `失败[${value.error_code}]${value.retryable ? '（可重试）' : ''}: ${value.error}` }]
  const out = value.output
  return [{ type: 'text', text: typeof out === 'string' ? out : JSON.stringify(out) }]
}

const ENVELOPE_SCHEMA = {
  type: 'object',
  properties: {
    ok: { type: 'boolean', description: '是否成功' },
    operation: { type: 'string', description: '工具名' },
    changed: { type: 'boolean', description: '是否修改了文档' },
    saved: { type: 'boolean', description: '是否已保存' },
    verified: { type: 'boolean', description: '结果是否经过校验' },
    warnings: { type: 'array', items: { type: 'string' }, description: '非致命提示' },
    output: { description: '工具结果载荷' },
    error_code: { type: 'string', description: '稳定错误码（CHANNEL_UNAVAILABLE / FILE_LOCKED / MISSING_PARAM …）' },
    error: { type: 'string', description: '人类可读错误' },
    retryable: { type: 'boolean', description: '是否可重试' },
    partial_changes: { type: 'boolean', description: '失败前是否已有部分修改残留' },
  },
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
  const operation = tool.name
  const meta = TOOL_META[operation] || {}
  const rawExecute = tool.execute
  ctx.tools.register({
    ...tool,
    execute: async (args) => finalize(operation, meta, args, await rawExecute(args)),
    parameters: { type: 'object', properties: props, required },
    output: { schema: ENVELOPE_SCHEMA, render: renderText },
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
  'wb, _opened = _resolve_wb(excel)',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  'try:',
  "    ws = wb.Worksheets(args['sheet']) if args.get('sheet') else wb.ActiveSheet",
  "    ws.Range(args['range']).Formula = args['formula']",
  "    output = json.dumps({'formula_set': '%s!%s = %s' % (ws.Name, args['range'], args['formula'])}, ensure_ascii=False)",
  '    if _opened: wb.Save()',
  'finally:',
  '    if _opened: wb.Close(False)',
])

// excel_recalc — 强制重算 + 取计算后活值（日期/NaN 清洗同 read_range）
const CODE_RECALC = comCode([
  'import datetime, math',
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  'wb, _opened = _resolve_wb(excel)',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  'try:',
  '    excel.Calculate()',
  '    def san(v):',
  '        if v is None: return None',
  '        if isinstance(v, (datetime.datetime, datetime.date)): return v.isoformat()',
  '        if isinstance(v, float) and not math.isfinite(v): return None',
  '        return v',
  '    def tolist(x):',
  '        if isinstance(x, (tuple, list)): return [tolist(i) for i in x]',
  '        return san(x)',
  "    if args.get('range'):",
  "        ws = wb.Worksheets(args['sheet']) if args.get('sheet') else wb.ActiveSheet",
  "        v = tolist(ws.Range(args['range']).Value)",
  "        output = json.dumps({'recalc': 'ok', 'range': args['range'], 'value': v}, ensure_ascii=False)",
  '    else:',
  "        output = json.dumps({'recalc': 'ok'}, ensure_ascii=False)",
  'finally:',
  '    if _opened: wb.Close(False)',
])

// excel_vba_run — 运行已有宏 / 注入执行
const CODE_VBA_RUN = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  'wb, _opened = _resolve_wb(excel)',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  'try:',
  "    r = excel.Run(args['macro'])",
  "    output = json.dumps({'vba_run': args['macro'], 'result': r}, ensure_ascii=False, default=str)",
  '    if _opened: wb.Save()',
  'finally:',
  '    if _opened: wb.Close(False)',
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
// 置可见：officemcp 的 Dispatch 实例默认不可见，新建的工作簿用户既看不见、也没有保存工具能落盘，
// 等于凭空多一个够不着的孤儿（退出时还会被收尾回收掉）。切成可见 = 交到用户手上的真实工作簿，
// 与 officemcp 自己的 Launch 行为一致；已有的可见实例上这一步是 no-op。
const CODE_NEW = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  'excel.Visible = True',
  'wb = excel.Workbooks.Add()',
  "sheets = [ws.Name for ws in wb.Worksheets]",
  "output = json.dumps({'new_workbook': True, 'sheets': sheets}, ensure_ascii=False)",
])

// excel_open — 打开已有工作簿，返回结构信息
const CODE_OPEN = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  "if not args.get('path'): raise Exception('缺少 path')",
  'wb, _opened = _resolve_wb(excel)',
  "if wb is None: raise Exception('无法打开工作簿: ' + args['path'])",
  'try:',
  '    sheets = [ws.Name for ws in wb.Worksheets]',
  "    output = json.dumps({'path': args['path'], 'active_sheet': wb.ActiveSheet.Name, 'sheets': sheets}, ensure_ascii=False)",
  'finally:',
  '    if _opened: wb.Close(False)',
])

// excel_read_range — 读区域活值（2D 数组；日期→iso，NaN/Inf→null）
const CODE_READ_RANGE = comCode([
  'import datetime, math',
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  'wb, _opened = _resolve_wb(excel)',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  'try:',
  '    def san(v):',
  '        if v is None: return None',
  '        if isinstance(v, (datetime.datetime, datetime.date)): return v.isoformat()',
  '        if isinstance(v, float) and not math.isfinite(v): return None',
  '        return v',
  '    def tolist(x):',
  '        if isinstance(x, (tuple, list)): return [tolist(i) for i in x]',
  '        return san(x)',
  "    ws = wb.Worksheets(args['sheet']) if args.get('sheet') else wb.ActiveSheet",
  "    output = json.dumps({'range': args['range'], 'value': tolist(ws.Range(args['range']).Value)}, ensure_ascii=False)",
  'finally:',
  '    if _opened: wb.Close(False)',
])

// excel_write_range — 写值（标量或 2D 数组；list→tuple 供 COM 赋值）
const CODE_WRITE_RANGE = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  'wb, _opened = _resolve_wb(excel)',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  'try:',
  "    ws = wb.Worksheets(args['sheet']) if args.get('sheet') else wb.ActiveSheet",
  '    def tocom(v):',
  '        if isinstance(v, list): return tuple(tocom(i) for i in v)',
  '        return v',
  '    val = tocom(args.get("value"))',
  '    ws.Range(args["range"]).Value = val',
  "    output = json.dumps({'range': args['range'], 'written': True}, ensure_ascii=False)",
  '    if _opened: wb.Save()',
  'finally:',
  '    if _opened: wb.Close(False)',
])

// word_open — 打开已有 Word 文档，返回结构信息（段落数 + 文本预览）
const CODE_WORD_OPEN = comCode([
  'word = Officer.Word',
  "if not word: raise Exception('Word 未安装/不可用')",
  'word.DisplayAlerts = 0',
  "if not args.get('path'): raise Exception('缺少 path')",
  'doc, _opened = _resolve_doc(word)',
  'try:',
  '    text = doc.Content.Text',
  "    output = json.dumps({'path': args['path'], 'paragraphs': doc.Paragraphs.Count, 'chars': len(text), 'preview': text[:300]}, ensure_ascii=False)",
  'finally:',
  '    if _opened: doc.Close(False)',
])

// word_edit — 全文查找替换（真实 Word Find.Execute + wdReplaceAll，替换只动文本、保留原文格式；
//   区别于整篇 Content.Text 覆盖——那种写法会把加粗/颜色/字体等排版全部抹掉）
const CODE_WORD_EDIT = comCode([
  'word = Officer.Word',
  "if not word: raise Exception('Word 未安装/不可用')",
  'word.DisplayAlerts = 0',
  'find_text = args.get("find", "")',
  "if not find_text: raise Exception('缺少 find（查找文本不能为空）')",
  'replace_text = args.get("replace", "")',
  'doc, _opened = _resolve_doc(word)',
  "if doc is None: raise Exception('无活动文档且未给 path')",
  'try:',
  '    # 先计数：纯 Find 查找不动文档（Forward + wdFindStop=0 防回绕死循环）',
  '    rng = doc.Content',
  '    rng.Find.ClearFormatting()',
  '    count = 0',
  '    while rng.Find.Execute(FindText=find_text, Forward=True, Wrap=0):',
  '        count += 1',
  '        rng.Collapse(0)',
  '    # 再替换：wdReplaceAll=2，只替换文本、保留原文格式。',
  '    # 注意必须用位置参数——实测 pywin32 对 Execute(FindText=..., ReplaceWith=..., Replace=...) 具名绑定不生效（返回 True 但文本不变），位置绑定才可靠',
  '    if count > 0:',
  '        doc.Content.Find.ClearFormatting()',
  '        doc.Content.Find.Execute(find_text, False, False, False, False, False, True, 0, False, replace_text, 2)',
  '        doc.Save()',
  "    output = json.dumps({'find': find_text, 'replace': replace_text, 'count': count, 'saved': count > 0}, ensure_ascii=False)",
  'finally:',
  '    if _opened: doc.Close(False)',
])

// excel_pivot_create — 创建透视表（PivotCache + PivotTable，行/列/值字段）
const CODE_PIVOT_CREATE = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  "if not args.get('range'): raise Exception('缺少 range（源数据区域，含表头）')",
  'wb, _opened = _resolve_wb(excel)',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  'try:',
  "    ws = wb.Worksheets(args['sheet']) if args.get('sheet') else wb.ActiveSheet",
  '    ws_dst = wb.Worksheets.Add()',
  '    n = 0',
  '    for w in wb.Worksheets:',
  '        n += len(list(w.PivotTables()))',
  "    tname = 'PivotTable%d' % (n + 1)",
  "    pc = wb.PivotCaches().Create(SourceType=1, SourceData=ws.Range(args['range']))", // SourceType=1 xlDatabase
  "    pt = pc.CreatePivotTable(TableDestination=ws_dst.Range('A1'), TableName=tname)",
  "    for f in (args.get('rows') or []):",
  '        pt.PivotFields(f).Orientation = 1', // xlRowField
  "    for f in (args.get('columns') or []):",
  '        pt.PivotFields(f).Orientation = 2', // xlColumnField
  "    for f in (args.get('values') or []):",
  "        pt.AddDataField(pt.PivotFields(f), f + '_求和', -4157)", // xlSum
  "    output = json.dumps({'pivot': 'ok', 'dst_sheet': ws_dst.Name, 'rows': args.get('rows'), 'columns': args.get('columns'), 'values': args.get('values')}, ensure_ascii=False)",
  '    if _opened: wb.Save()',
  'finally:',
  '    if _opened: wb.Close(False)',
])

// excel_pivot_refresh — 刷新工作簿内所有透视表
const CODE_PIVOT_REFRESH = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  'wb, _opened = _resolve_wb(excel)',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  'try:',
  '    count = 0',
  '    for ws in wb.Worksheets:',
  '        for pt in ws.PivotTables():',
  '            pt.RefreshTable()',
  '            count += 1',
  "    output = json.dumps({'refreshed': count}, ensure_ascii=False)",
  '    if _opened: wb.Save()',
  'finally:',
  '    if _opened: wb.Close(False)',
])

// excel_journal_post — 写会计分录到账簿 + 借贷平衡校验（错则标红）
const CODE_JOURNAL_POST = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  'wb, _opened = _resolve_wb(excel)',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  'try:',
  "    ws = wb.Worksheets(args['sheet']) if args.get('sheet') else wb.ActiveSheet",
  '    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row', // xlUp 找 A 列最后一行
  "    if last == 1 and str(ws.Cells(1,1).Value or '').strip() == '':",
  "        ws.Cells(1,1).Value = '日期'",
  "        ws.Cells(1,2).Value = '摘要'",
  "        ws.Cells(1,3).Value = '科目'",
  "        ws.Cells(1,4).Value = '借方'",
  "        ws.Cells(1,5).Value = '贷方'",
  '        start = 2',
  '    else:',
  '        start = last + 1',
  '    dr = 0.0; cr = 0.0',
  "    entries = args.get('entries') or []",
  '    for i, e in enumerate(entries):',
  '        r = start + i',
  "        ws.Cells(r,1).Value = e.get('date','')",
  "        ws.Cells(r,2).Value = e.get('desc','')",
  "        ws.Cells(r,3).Value = e.get('account','')",
  '        d = float(e.get("debit") or 0); c = float(e.get("credit") or 0)',
  '        ws.Cells(r,4).Value = d',
  '        ws.Cells(r,5).Value = c',
  '        dr += d; cr += c',
  '    balanced = abs(dr - cr) < 1e-9',
  '    if not balanced and entries:',
  '        ws.Range(ws.Cells(start,1), ws.Cells(start+len(entries)-1,5)).Font.Color = 255', // 红色标错
  "    output = json.dumps({'posted': len(entries), 'debit_total': dr, 'credit_total': cr, 'balanced': balanced}, ensure_ascii=False)",
  '    if _opened: wb.Save()',
  'finally:',
  '    if _opened: wb.Close(False)',
])

// excel_ledger_gen — 日记账 → 科目总账（按科目聚合借/贷，余额=借-贷公式）
const CODE_LEDGER_GEN = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  'wb, _opened = _resolve_wb(excel)',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  'try:',
  "    ws = wb.Worksheets(args['journal_sheet']) if args.get('journal_sheet') else wb.ActiveSheet",
  '    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row',
  "    if last < 2: raise Exception('日记账为空（无数据行）')",
  '    accounts = {}',
  '    for r in range(2, last+1):',
  '        acc = ws.Cells(r,3).Value',
  '        if acc is None: continue',
  '        d = float(ws.Cells(r,4).Value or 0); c = float(ws.Cells(r,5).Value or 0)',
  '        if acc not in accounts: accounts[acc] = [0.0, 0.0]',
  '        accounts[acc][0] += d; accounts[acc][1] += c',
  "    ws_out = wb.Worksheets(args['output_sheet']) if args.get('output_sheet') else wb.Worksheets.Add()",
  "    ws_out.Cells(1,1).Value = '科目'; ws_out.Cells(1,2).Value = '借方合计'; ws_out.Cells(1,3).Value = '贷方合计'; ws_out.Cells(1,4).Value = '余额'",
  '    r = 2',
  '    for acc in sorted(accounts.keys()):',
  '        d, c = accounts[acc]',
  '        ws_out.Cells(r,1).Value = acc',
  '        ws_out.Cells(r,2).Value = d',
  '        ws_out.Cells(r,3).Value = c',
  "        ws_out.Cells(r,4).Formula = '=B%d-C%d' % (r, r)",
  '        r += 1',
  "    output = json.dumps({'accounts': len(accounts), 'output_sheet': ws_out.Name}, ensure_ascii=False)",
  '    if _opened: wb.Save()',
  'finally:',
  '    if _opened: wb.Close(False)',
])

// ── 插件入口 ─────────────────────────────────────────────────

// 导出内部函数供冒烟测试（DSH 只用 name/inject/apply）
export { findPython, ensureServer, runPython, finalize, classifyError, cleanupOwnedApps, runningOfficeImages, TOOL_META }

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
      value: {
        required: true,
        description: '写入的值：标量，或与区域形状一致的 2D 数组',
        oneOf: [
          { type: 'string' },
          { type: 'number' },
          { type: 'boolean' },
          { type: 'array' },
        ],
      },
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
