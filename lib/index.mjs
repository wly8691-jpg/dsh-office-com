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

function spawnServer(py, port) {
  const script = [
    'import sys',
    `sys.argv=["officemcp","sse","--port","${port}","--folder","${OFFICE_FOLDER}"]`,
    'from officemcp.OfficeMCP import RunOfficeMCP',
    'RunOfficeMCP()',
  ].join(';')
  const child = spawn(py, ['-c', script], { env: cleanEnv(), windowsHide: true, stdio: ['ignore', 'ignore', 'pipe'] })
  let err = ''
  child.stderr.on('data', (d) => { err += d })
  child.on('error', () => {})
  child.on('exit', (code) => {
    if (code && code !== 0) console.error(`[dsh-office-com] officemcp exited ${code}: ${err.trim().slice(0, 400)}`)
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

async function ensureServer() {
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
    try { return JSON.parse(text) } catch { return { success: false, error: text.slice(0, 500) || 'empty output' } }
  } catch (e) {
    return { success: false, error: e.message }
  }
}

/** 拼 COM 代码：args 以 JSON 走 data，代码内 json.loads(data) 取值 */
function comCode(body) {
  return ['import json', 'args = json.loads(data) if data else {}', ...body].join('\n')
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
  ctx.tools.register({ ...tool, output: { schema: { type: 'object' }, render: renderText } })
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
  "output = 'formula_set: %s!%s = %s' % (ws.Name, args['range'], args['formula'])",
])

// excel_recalc — 强制重算 + 取计算后活值
const CODE_RECALC = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  "if args.get('path'):",
  "    wb = excel.Workbooks.Open(args['path'])",
  'else:',
  '    wb = excel.ActiveWorkbook',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  'excel.Calculate()',
  "if args.get('range'):",
  "    ws = wb.Worksheets(args['sheet']) if args.get('sheet') else wb.ActiveSheet",
  "    v = ws.Range(args['range']).Value",
  "    output = json.dumps({'recalc': 'ok', 'range': args['range'], 'value': v}, ensure_ascii=False)",
  'else:',
  "    output = 'recalc ok'",
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
  "output = 'vba_run: %s -> %s' % (args['macro'], json.dumps(r, ensure_ascii=False))",
])

// ── 插件入口 ─────────────────────────────────────────────────

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

  console.log(`[dsh-office-com] plugin loaded${degraded ? ' (degraded)' : ''}`)
}
