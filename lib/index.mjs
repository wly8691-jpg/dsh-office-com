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

// 前导比几个辅助函数长得多，用模板串写比逐行引号数组可读（Python 侧一律单引号，与反引号不冲突）。
// 注意 `\\\\` 在模板串里落成 `\\`，正是 Python 源码表示单个反斜杠所需的写法。
const PREAMBLE = `import json, pythoncom, datetime
pythoncom.CoInitialize()
args = json.loads(data) if data else {}
if args.get('path'): args['path'] = args['path'].replace('/', '\\\\')

# ── 控制面 ────────────────────────────────────────────────────
# mode / save / close / overwrite / confirm 由 JS 侧 resolveCtl 解析一次后随 args._ctl 下发。
# 这里只消费、不推断——两处各推一遍模式必然漂移，而漂移正好砸在工单 §五「失败状态可解释率」上。

def _ctl():
    c = args.get('_ctl')
    if isinstance(c, dict): return c
    # 没接线时退回今天的语义（有 path 就存就关），退化成旧行为而不是**错**模式
    hp = bool(args.get('path'))
    return {'preview': False, 'want_save': hp, 'want_close': hp}

_ctx = {'wb': None, 'app': None, 'opened': False, 'existing_open': False,
        'changed': False, 'saved': False, 'closed': False, 'left_open': False,
        'finished': False, 'error': None, 'notes': []}

def _find_open(coll, path):
    # Excel/Word 的 Open 对已打开的同一个文件返回**同一个对象**。不先找一遍就会把它当成
    # "自己开的"在收尾时 Close，关掉用户正在编辑的窗口并丢弃其未保存修改（工单 §四 重点保护）。
    try:
        want = path.replace('/', '\\\\').lower()
        for d in coll:
            try:
                if str(d.FullName).lower() == want: return d
            except Exception: pass
    except Exception: pass
    return None

def _resolve_wb(excel):
    # 给 path 就 Open（owned=True），否则用活动工作簿（owned=False）。只有自己 Open 的才收尾关闭。
    _ctx['app'] = excel
    if args.get('path'):
        hit = _find_open(excel.Workbooks, args['path'])
        if hit is not None:
            _ctx['existing_open'] = True
            _ctx['wb'] = hit
            return hit, False
        wb = excel.Workbooks.Open(args['path'])
        _ctx['wb'] = wb; _ctx['opened'] = True
        return wb, True
    wb = excel.ActiveWorkbook
    _ctx['wb'] = wb
    return wb, False

def _resolve_doc(word):
    _ctx['app'] = word
    if args.get('path'):
        hit = _find_open(word.Documents, args['path'])
        if hit is not None:
            _ctx['existing_open'] = True
            _ctx['wb'] = hit
            return hit, False
        doc = word.Documents.Open(args['path'])
        _ctx['wb'] = doc; _ctx['opened'] = True
        return doc, True
    doc = word.ActiveDocument
    _ctx['wb'] = doc
    return doc, False

def _mark_changed():
    # 每处工具在**第一句变更语句之前**调一次，失败时才说得清有没有残留改动。
    _ctx['changed'] = True

def _norm_date(v):
    # Excel 会把 '2026-01-05' 这类字符串**存成真日期**，读回来是 datetime 而不是原字符串。
    # 拿它做批次签名时只留日期部分，否则"看起来一样"的两批永远对不上（幂等静默失效）。
    if isinstance(v, (datetime.datetime, datetime.date)): return v.strftime('%Y-%m-%d')
    s = str(v or '').strip()
    if 'T' in s: return s.split('T')[0]
    if ' ' in s: return s.split(' ')[0]
    return s

def _addr(rng):
    # Excel 的 Range.Address 是**带参属性**，pywin32 晚绑定下取到的直接就是字符串，
    # 再去调用它必然 'str' object is not callable —— 所以只读取、不调用。
    try: return str(rng.Address)
    except Exception: return None

def _flat(x):
    # COM 的 Range.Value：单元格是标量，多单元格是嵌套 tuple。展平后好数非空格。
    if isinstance(x, (tuple, list)):
        out = []
        for i in x: out.extend(_flat(i))
        return out
    return [x]

def _shape(v):
    # 输入值的形状描述，用来和真实区域形状对账（形状不匹配是写坏表格的常见起因）。
    if isinstance(v, list) and v and isinstance(v[0], list): return '%dx%d' % (len(v), len(v[0]))
    if isinstance(v, list): return '1x%d' % len(v)
    return 'scalar'

def _partial():
    # 失败后工作簿是否仍停在被修改状态——只有 Python 知道，JS 只能靠 args 猜。
    # 自己开的那些收尾时 Close(False) 会丢弃，不留残迹；附着/用户已在用的，改动留在那。
    return bool(_ctx['changed'] and (_ctx['left_open'] or not _ctx['opened']))

def _close_wb():
    try:
        _ctx['wb'].Close(False)
        _ctx['closed'] = True
    except Exception as e:
        _ctx['left_open'] = True
        _ctx['notes'].append('close_failed')
    _ctx['wb'] = None

def _finish(ok):
    # 收尾只在**成功路径**上保存。放进 finally 会把"改了一半"的修改落盘，与今天
    # Close(False) 的回滚语义正好相反——工单 §五 的「部分修改状态可识别率」量的是这个。
    if _ctx['finished']: return
    _ctx['finished'] = True
    if _ctx['wb'] is None: return
    c = _ctl()
    if c.get('preview'):
        if _ctx['opened']: _close_wb()      # preview：只关自己开的那份，永不落盘
        return
    # 落盘许可：自己开的那份按 want_save 存；attached 下只有**显式** save:true 才存
    # （默认代存会把用户在编辑的未保存内容一并落盘，那不是他要的）。
    # 文件本就在用户手里（existing_open）一律不代存 —— 同理。
    may_save = _ctx['opened'] or (c.get('save_requested') and not _ctx['existing_open'])
    # 没有改动就不落盘：省一次无谓的写盘与 mtime 变动（幂等跳过的路径会走到这）
    if ok and c.get('want_save') and may_save and _ctx['changed']:
        try:
            _ctx['wb'].Save()
            _ctx['saved'] = True
        except Exception as e:
            _ctx['error'] = '[SAVE_FAILED] 保存失败: %s' % e
    if not _ctx['opened']:
        # attached，或文件本就在用户手里：不主动关
        if _ctx['existing_open']: _ctx['notes'].append('existing_open_no_save')
        return
    if c.get('want_close', True):
        _close_wb()
    else:
        _ctx['left_open'] = True
        _ctx['notes'].append('leave_open')

def _plan(extra):
    # 必带工作簿身份：工单 §三 要求证明碰的是真实运行中的那个 Excel/Word。
    wb = _ctx['wb']
    p = {'preview': True, 'tool': _ctl().get('tool'), 'mode': _ctl().get('mode'),
         'workbook': (wb.Name if wb is not None else None),
         'path': (wb.FullName if wb is not None else args.get('path')),
         'changed': False, 'saved': False}
    try: p['app_visible'] = bool(_ctx['app'].Visible)
    except Exception: p['app_visible'] = None
    p.update(extra)
    return p

class _Preview(Exception):
    def __init__(self, plan):
        Exception.__init__(self, 'preview')
        self.plan = plan

def _emit_preview(plan):
    # 放在第一句变更语句**之前**。preview 的职责是报告 would_clobber，
    # 不是报 OVERWRITE_NOT_CONFIRMED 失败——顺序反了 Agent 就拿不到预演结果。
    if _ctl().get('preview'): raise _Preview(plan)

def _fail(msg, **meta):
    # Python 侧上报稳定错误码的**唯一**出口：COM 报错文本随语言/版本变化，模糊匹配不可靠。
    raise Exception('%s [META]%s' % (msg, json.dumps(meta, ensure_ascii=False)))

def _observed(output):
    # 观测结果并进 output，回给 JS 的信封。声明是意图，观测是事实。
    try:
        o = json.loads(output)
    except Exception:
        return output
    if not isinstance(o, dict): return output
    o['changed'] = bool(_ctx['changed'])
    o['saved'] = bool(_ctx['saved'])
    if _ctx['existing_open']: o['attached_existing_open'] = True
    if _ctx['closed']: o['closed'] = True
    if _ctx['left_open']: o['left_open'] = True
    if _ctx['notes']: o['notes'] = _ctx['notes']
    return json.dumps(o, ensure_ascii=False, default=str)
`

// 收尾：异常路径要走 _finish（否则自己开的工作簿不会被关，泄漏文件锁与进程）；
// 成功路径的 Save 失败由 _ctx['error'] 兜住，同样走异常路径回给信封。
const EPILOGUE = [
  'except _Preview as _p:',
  '    _finish(False)',
  '    output = json.dumps(_p.plan, ensure_ascii=False, default=str)',
  'except Exception as _e:',
  '    _finish(False)',
  "    raise Exception('%s [META]%s' % (_e, json.dumps({'partial_changes': _partial()}, ensure_ascii=False)))",
  "if _ctx['error']:",
  "    raise Exception('%s [META]%s' % (_ctx['error'], json.dumps({'partial_changes': _partial()}, ensure_ascii=False)))",
  'output = _observed(output)',
]

/** 拼 COM 代码：args 以 JSON 走 data，代码内 json.loads(data) 取值。CoInitialize 是 COM Dispatch 的线程前置（否则 -2147221008）。
 *  path 统一归一化为反斜杠——实测 Excel SaveAs/Open 对正斜杠路径解析异常（会把 D:/x 拼成乱路径）。
 *  body 整体缩进一层放进 try：各工具不再自带 try/finally，收尾统一由 _finish 负责。 */
function comCode(body) {
  return [PREAMBLE, 'try:', ...body.map((l) => (l ? '    ' + l : l)), ...EPILOGUE].join('\n')
}

// ── 统一结果协议 v0.2 ────────────────────────────────────────
// 所有工具返回同一信封，Agent 只需读 ok / error_code / retryable 就能分支，
// 不必解析每个工具各自不同的返回形状。
//   成功：{ ok, operation, mode, preview, changed, saved, verified, warnings, output }
//   失败：{ ok, operation, mode, preview, error_code, error, retryable, partial_changes }

const TOOL_META = {
  office_apps:         { changed: false, saved: false, verified: false },
  office_launch:       { changed: false, saved: false, verified: false },
  excel_new:           { changed: false, saved: false, verified: false },
  excel_open:          { changed: false, saved: false, verified: false },
  excel_read_range:    { changed: false, saved: false, verified: false },
  excel_recalc:        { changed: false, saved: false, verified: true  }, // 重算后读回活值 = 已验证
  excel_formula_set:   { changed: true,  saved: true,  verified: false },
  excel_write_range:   { changed: true,  saved: true,  verified: false },
  excel_vba_run:       { changed: true,  saved: true,  verified: false },
  excel_pivot_create:  { changed: true,  saved: true,  verified: false },
  excel_pivot_refresh: { changed: true,  saved: true,  verified: false },
  excel_journal_post:  { changed: true,  saved: true,  verified: true  }, // 借贷平衡校验 = 已验证
  excel_ledger_gen:    { changed: true,  saved: true,  verified: false },
  word_open:           { changed: false, saved: false, verified: false },
  word_edit:           { changed: true,  saved: true,  verified: false },
}

// ── 安全动作模式 v0.3 ────────────────────────────────────────
// preview：只分析并返回拟执行变更，不修改文件（真经 COM 打开看，只读、不存、看完关）
// managed：打开 → 修改 → 校验 → 保存 → 关闭（缺省带 path 时的行为）
// attached：操作当前运行实例，不主动关闭、不自动覆盖（缺省不带 path 时的行为）
//
// 控制面在 JS 侧解析**一次**，随 args._ctl 下发给 Python。两处各推一遍模式必然漂移，
// 而漂移正好砸在工单 §五「失败状态可解释率」上——所以 Python 只消费、不推断。

const MODES = ['preview', 'managed', 'attached']

/** 会改动文档的工具（其余工具接受 mode 但不产生 dry-run 差异）。 */
const MUTATING_TOOLS = new Set([
  'excel_formula_set', 'excel_write_range', 'excel_vba_run',
  'excel_pivot_create', 'excel_pivot_refresh',
  'excel_journal_post', 'excel_ledger_gen', 'word_edit',
])

/** mode / save / close / overwrite / confirm 的共享参数块——描述只写一份，防各工具间漂移。
 *  刻意不写 required（缺省即非必填）：reg() 会把各属性的 required 摘掉收进顶层数组，
 *  共享对象被就地改写会互相串味。 */
const MODE_PARAMS = {
  mode: {
    type: 'string', enum: MODES,
    description: '执行模式：preview 只预演不落笔 / managed 打开-改-存-关 / attached 改当前运行实例且不关不存。缺省按有无 path 推断（有=managed，无=attached）',
  },
  save: { type: 'boolean', description: '是否落盘。managed 缺省 true、attached 缺省 false。高风险参数，显式声明才生效' },
  close: { type: 'boolean', description: '是否关闭。managed 缺省 true（自己开的那份）；attached 下无效——用户的工作簿永不代关' },
  overwrite: { type: 'boolean', description: '是否允许覆盖既有内容（如已存在的总账工作表）。涉及覆盖时必填 true，否则返回 OVERWRITE_NOT_CONFIRMED' },
  confirm: { type: 'boolean', description: '高风险动作确认位（如运行 VBA 宏）。缺省 false = 拒绝执行' },
}

/** 参数 → 控制面。非法 mode 不抛异常，交由 finalize 出稳定错误码（信封要一致）。 */
function resolveCtl(operation, args) {
  const a = args || {}
  const hasPath = a.path != null && a.path !== ''
  const mode = a.mode ?? (hasPath ? 'managed' : 'attached')
  const preview = mode === 'preview'
  const valid = MODES.includes(mode)
  return {
    tool: operation,
    mode,
    valid,
    preview,
    mutating: MUTATING_TOOLS.has(operation),
    // preview 永不落盘；关只关自己开的那份（Python 侧按 owned 判定）
    // 默认落盘只给会改动文档的工具：读工具（read_range / recalc / open / word_open）也走 managed
    // 打开文件，但今天就不该存盘——别让模式改动把它们顺手写坏。
    want_save: preview ? false : (a.save != null ? !!a.save : (mode === 'managed' && MUTATING_TOOLS.has(operation))),
    want_close: preview ? true : (a.close != null ? !!a.close : mode === 'managed'),
    overwrite: !!a.overwrite,
    confirm: !!a.confirm,
    save_requested: a.save != null,
    close_requested: a.close != null,
  }
}

/** 原始 args + 控制面 → 走 data 通道的 JSON 串。 */
function argsData(args, ctl) {
  return JSON.stringify({ ...(args || {}), _ctl: ctl })
}

/** Python 侧用 [META]{...} 尾缀回传失败元数据（桥在 isError 时丢弃 output，错误文本是唯一通道）。 */
function parseErrMeta(error) {
  const s = String(error || '')
  const m = /\[META\](\{.*\})\s*$/.exec(s)
  if (!m) return { text: s, meta: {} }
  try {
    return { text: s.slice(0, m.index).trim(), meta: JSON.parse(m[1]) }
  } catch {
    return { text: s, meta: {} }
  }
}

/** Python 带 [CODE] 前缀上报的稳定码——COM 报错文本随语言/版本变，模糊匹配不可靠。 */
const PREFIX_CODES = {
  MODE_INVALID: false,
  OVERWRITE_NOT_CONFIRMED: false,
  RISKY_OP_NOT_CONFIRMED: false,
  SAVE_FAILED: true,      // 文件被占/只读/磁盘满多为瞬时
  VBA_FAILED: false,
  PARTIAL_WRITE: false,
  SOURCE_RANGE_INVALID: false,
  MISSING_PARAM: false,
  NO_ACTIVE_DOCUMENT: false,
  APP_UNAVAILABLE: false,
  EMPTY_SOURCE: false,
  OPEN_FAILED: false,
}

/** 把 Python 抛出的错误文本映射到稳定 error_code + 可重试性。 */
function classifyError(error) {
  const e = String(error || '')
  if (/SSE not connected|MCP call timeout|fetch failed|ECONNREFUSED|socket hang up|Failed to fetch|network/i.test(e)) {
    return { code: 'CHANNEL_UNAVAILABLE', retryable: true } // 通道断了，重连可重试
  }
  // 前缀码排在上面的通道分支**之后**：SSE 超时文本里可能出现任何字样，不能被误分类。
  const p = /\[([A-Z][A-Z0-9_]+)\]/.exec(e)
  if (p && p[1] in PREFIX_CODES) return { code: p[1], retryable: PREFIX_CODES[p[1]] }
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

/** 原始 {ok, output} / {ok:false, error} → 统一信封。
 *  ctl 可选：缺省自行解析，这样外部按旧签名调用（4 参）照旧可用。
 *  字段取值**观测优先**——Python 侧知道真实发生了什么（改没改、存没存、还开着没），
 *  JS 只能靠 args 猜。两者冲突时一律信观测：信封绝不能宣称一次 COM 侧已知失败的保存。 */
function finalize(operation, meta, args, r, ctl) {
  const c = ctl || resolveCtl(operation, args)
  if (r && r.ok === true) {
    const out = r.output
    const observed = out && typeof out === 'object' ? out : {}
    const warnings = []
    let verified = meta.verified ?? false
    // 会计特例：借贷不平衡 = 成功但带 warning（verified 置 false，不是错误是提示）
    if (observed.balanced === false) {
      verified = false
      warnings.push('借贷不平衡：借方合计 ≠ 贷方合计')
    }
    const changed = observed.changed ?? (c.preview ? false : (meta.changed ?? false))
    let saved = observed.saved ?? ((meta.saved ?? false) && c.want_save)
    if (observed.attached_existing_open === true) {
      // U0：目标文件已在用户手里开着 → 只改了内存，没关也没落盘。信封必须说实话。
      saved = false
      warnings.push('目标文件已被用户打开，按附着处理：已改内存，未落盘、也未关闭')
    }
    if (c.preview) warnings.push('preview：未做任何修改')
    else if (changed && c.mode === 'attached' && !saved) warnings.push('attached 模式：已修改当前实例，未落盘（要落盘请显式 save:true）')
    if (Array.isArray(observed.notes)) {
      if (observed.notes.includes('leave_open')) warnings.push('close:false：遗留已打开的工作簿（文件锁/进程泄漏）')
    }
    return {
      ok: true,
      operation,
      mode: c.mode,
      preview: c.preview,
      changed,
      saved,
      verified,
      warnings,
      output: out,
    }
  }
  const { text, meta: errMeta } = parseErrMeta((r && r.error) || 'unknown error')
  const cls = classifyError(text)
  return {
    ok: false,
    operation,
    mode: c.mode,
    preview: c.preview,
    error_code: cls.code,
    error: text,
    retryable: cls.retryable,
    // 观测优先：只有 Python 知道失败后工作簿是否仍停在被修改状态（attached 改到一半、
    // 或 managed 的 Save 失败后 Close 也失败）。拿不到观测才按模式兜底。
    partial_changes: errMeta.partial_changes ?? (!!meta.changed && c.mode !== 'managed' && !c.preview),
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
    mode: { type: 'string', description: '实际执行的模式：preview / managed / attached' },
    preview: { type: 'boolean', description: '是否为预演（true 时 changed/saved 恒为 false）' },
    changed: { type: 'boolean', description: '是否修改了文档' },
    saved: { type: 'boolean', description: '是否已保存（观测所得，不是意图）' },
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
    execute: async (args) => {
      // 控制面只在这里解析一次，同时喂 finalize（信封口径）和 rawExecute（下发给 Python）
      const ctl = resolveCtl(operation, args)
      if (!ctl.valid) {
        return finalize(operation, meta, args, {
          ok: false,
          error: `[MODE_INVALID] 未知 mode: ${JSON.stringify(args?.mode)}（应为 preview / managed / attached）`,
        }, ctl)
      }
      return finalize(operation, meta, args, await rawExecute(args, ctl), ctl)
    },
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
  "ws = wb.Worksheets(args['sheet']) if args.get('sheet') else wb.ActiveSheet",
  'rng = ws.Range(args["range"])',
  'try: _before = rng.Formula',
  'except Exception: _before = None',
  "_emit_preview(_plan({'target': {'sheet': ws.Name, 'address': _addr(rng), 'cells': rng.Cells.Count}, 'before': _before, 'after': args['formula'], 'would_change': _before != args['formula']}))",
  '_mark_changed()',
  "rng.Formula = args['formula']",
  "output = json.dumps({'formula_set': '%s!%s = %s' % (ws.Name, args['range'], args['formula'])}, ensure_ascii=False)",
  '_finish(True)',
])

// excel_recalc — 强制重算 + 取计算后活值（日期/NaN 清洗同 read_range）
const CODE_RECALC = comCode([
  'import datetime, math',
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  'wb, _opened = _resolve_wb(excel)',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  'def san(v):',
  '    if v is None: return None',
  '    if isinstance(v, (datetime.datetime, datetime.date)): return v.isoformat()',
  '    if isinstance(v, float) and not math.isfinite(v): return None',
  '    return v',
  'def tolist(x):',
  '    if isinstance(x, (tuple, list)): return [tolist(i) for i in x]',
  '    return san(x)',
  'excel.Calculate()',
  "if args.get('range'):",
  "    ws = wb.Worksheets(args['sheet']) if args.get('sheet') else wb.ActiveSheet",
  "    v = tolist(ws.Range(args['range']).Value)",
  "    output = json.dumps({'recalc': 'ok', 'range': args['range'], 'value': v}, ensure_ascii=False)",
  'else:',
  "    output = json.dumps({'recalc': 'ok'}, ensure_ascii=False)",
  '_finish(True)',
])

// excel_vba_run — 运行已有宏 / 注入执行
const CODE_VBA_RUN = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  'wb, _opened = _resolve_wb(excel)',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  'try:',
  '    _mods = [c.Name for c in wb.VBProject.VBComponents]',
  'except Exception:',
  '    # 未勾选"信任对 VBA 工程对象模型的访问"时读不到——报 null，别把预演本身搞失败',
  '    _mods = None',
  "_emit_preview(_plan({'macro': args['macro'], 'module_candidates': _mods, 'requires_confirm': True, 'confirm_present': bool(_ctl().get('confirm')), 'note': '宏的具体效果无法预演，preview 只说明将要运行哪个宏、以及确认位是否就位'}))",
  '# 高危门控放在 preview **之后**：preview 的职责是告诉 Agent 还差什么，而不是直接把它挡回去',
  "if not _ctl().get('confirm'):",
  "    raise Exception('[RISKY_OP_NOT_CONFIRMED] 运行 VBA 宏是高危动作（宏可以改动任意内容），请显式 confirm:true 再执行')",
  '_mark_changed()',
  'try:',
  "    r = excel.Run(args['macro'])",
  'except Exception as _ve:',
  '    # COM 的宏报错文本随语言/版本变化（"无法运行宏" / "Cannot run the macro"），不进模糊匹配梯队',
  "    raise Exception('[VBA_FAILED] 宏执行失败 %s: %s' % (args['macro'], _ve))",
  "output = json.dumps({'vba_run': args['macro'], 'result': r}, ensure_ascii=False, default=str)",
  '_finish(True)',
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
  'sheets = [ws.Name for ws in wb.Worksheets]',
  "output = json.dumps({'path': args['path'], 'active_sheet': wb.ActiveSheet.Name, 'sheets': sheets}, ensure_ascii=False)",
  '_finish(True)',
])

// excel_read_range — 读区域活值（2D 数组；日期→iso，NaN/Inf→null）
const CODE_READ_RANGE = comCode([
  'import datetime, math',
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  'wb, _opened = _resolve_wb(excel)',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  'def san(v):',
  '    if v is None: return None',
  '    if isinstance(v, (datetime.datetime, datetime.date)): return v.isoformat()',
  '    if isinstance(v, float) and not math.isfinite(v): return None',
  '    return v',
  'def tolist(x):',
  '    if isinstance(x, (tuple, list)): return [tolist(i) for i in x]',
  '    return san(x)',
  "ws = wb.Worksheets(args['sheet']) if args.get('sheet') else wb.ActiveSheet",
  "output = json.dumps({'range': args['range'], 'value': tolist(ws.Range(args['range']).Value)}, ensure_ascii=False)",
  '_finish(True)',
])

// excel_write_range — 写值（标量或 2D 数组；list→tuple 供 COM 赋值）
const CODE_WRITE_RANGE = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  'wb, _opened = _resolve_wb(excel)',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  "ws = wb.Worksheets(args['sheet']) if args.get('sheet') else wb.ActiveSheet",
  'rng = ws.Range(args["range"])',
  'def tocom(v):',
  '    if isinstance(v, list): return tuple(tocom(i) for i in v)',
  '    return v',
  'val = tocom(args.get("value"))',
  '_cells = [v for v in _flat(rng.Value) if v not in (None, "")]',
  "_emit_preview(_plan({'target': {'sheet': ws.Name, 'address': _addr(rng), 'shape': '%dx%d' % (rng.Rows.Count, rng.Columns.Count)}, 'input_shape': _shape(args.get('value')), 'overwrite_non_empty': len(_cells), 'sample_before': [str(v)[:40] for v in _cells[:5]]}))",
  '_mark_changed()',
  'rng.Value = val',
  "output = json.dumps({'range': args['range'], 'written': True}, ensure_ascii=False)",
  '_finish(True)',
])

// word_open — 打开已有 Word 文档，返回结构信息（段落数 + 文本预览）
const CODE_WORD_OPEN = comCode([
  'word = Officer.Word',
  "if not word: raise Exception('Word 未安装/不可用')",
  'word.DisplayAlerts = 0',
  "if not args.get('path'): raise Exception('缺少 path')",
  'doc, _opened = _resolve_doc(word)',
  'text = doc.Content.Text',
  "output = json.dumps({'path': args['path'], 'paragraphs': doc.Paragraphs.Count, 'chars': len(text), 'preview': text[:300]}, ensure_ascii=False)",
  '_finish(True)',
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
  '# 先计数：纯 Find 查找不动文档（Forward + wdFindStop=0 防回绕死循环）',
  'rng = doc.Content',
  'rng.Find.ClearFormatting()',
  'count = 0',
  '_first = None',
  'while rng.Find.Execute(FindText=find_text, Forward=True, Wrap=0):',
  '    count += 1',
  '    if count == 1:',
  '        try: _first = str(rng.Paragraphs(1).Range.Text)[:200]',
  '        except Exception: _first = None',
  '    rng.Collapse(0)',
  '# 计数循环本身不碰文档，所以 preview 是白送的；第二次跑同一 find 会得到 count=0 = 天然幂等',
  "_emit_preview(_plan({'doc': doc.Name, 'find': find_text, 'replace': replace_text, 'count': count, 'would_change': count > 0, 'first_hit_context': _first}))",
  '# 再替换：wdReplaceAll=2，只替换文本、保留原文格式。',
  '# 注意必须用位置参数——实测 pywin32 对 Execute(FindText=..., ReplaceWith=..., Replace=...) 具名绑定不生效（返回 True 但文本不变），位置绑定才可靠',
  'if count > 0:',
  '    _mark_changed()',
  '    doc.Content.Find.ClearFormatting()',
  '    doc.Content.Find.Execute(find_text, False, False, False, False, False, True, 0, False, replace_text, 2)',
  "output = json.dumps({'find': find_text, 'replace': replace_text, 'count': count}, ensure_ascii=False)",
  '_finish(True)',
])

// excel_pivot_create — 创建或复用它已有的透视表（PivotCache + PivotTable，行/列/值字段）
// 重跑安全：字段校验提到任何变更之前、目标表与表名都取确定性名字，第二次跑命中同名表就复用刷新，
// 而不是再建一张表 + 再建一个工作表。
const CODE_PIVOT_CREATE = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  "if not args.get('range'): raise Exception('缺少 range（源数据区域，含表头）')",
  'wb, _opened = _resolve_wb(excel)',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  "ws = wb.Worksheets(args['sheet']) if args.get('sheet') else wb.ActiveSheet",
  'src = ws.Range(args["range"])',
  // 旧代码是 `Worksheets.Add()` 之后才在 PivotFields 上炸，attached 模式会留下孤儿工作表；
  // 而且表名按现存透视表计数（PivotTable1/2/3…），重跑必然堆积。
  '# 字段校验提到任何变更之前',
  "_fields = [str(h) for h in _flat(src.Rows(1).Value) if h not in (None, '')]",
  "_want = list(args.get('rows') or []) + list(args.get('columns') or []) + list(args.get('values') or [])",
  '_missing = [f for f in _want if f not in _fields]',
  "if _missing: raise Exception('[SOURCE_RANGE_INVALID] 源数据区域缺少字段：%s（表头实有：%s）' % ('、'.join(_missing), '、'.join(_fields)))",
  '_dst_name = args.get("output_sheet") or "透视表"',
  "_tname = ('PT_' + ws.Name + '_' + args['range']).replace(' ', '_').replace('$', '').replace(':', '_').replace('/', '_')[:200]",
  'try: _ws_dst = wb.Worksheets(_dst_name)',
  'except Exception: _ws_dst = None',
  '_exist = None',
  'for _w in wb.Worksheets:',
  '    for _pt in _w.PivotTables():',
  '        if _pt.Name == _tname: _exist = (_w, _pt)',
  '_n_dst_pt = len(list(_ws_dst.PivotTables())) if _ws_dst is not None else 0',
  '_same_as_src = (_dst_name == ws.Name)',
  "_emit_preview(_plan({'source': {'sheet': ws.Name, 'range': args['range'], 'header_fields': _fields}, 'destination': {'sheet': _dst_name, 'cell': 'A1', 'exists': _ws_dst is not None, 'pivot_tables': _n_dst_pt}, 'table_name': _tname, 'would_reuse': _exist is not None, 'would_write_over_source': bool(_same_as_src), 'would_clobber': bool(_n_dst_pt > 0 and _exist is None)}))",
  '# 预演先出（它要**报告** would_clobber，不是替 Agent 做决定），守卫放后面',
  "if _same_as_src and not _ctl().get('overwrite'):",
  "    raise Exception('[OVERWRITE_NOT_CONFIRMED] 目标工作表与源数据表同名（%s），会在源数据上建透视表；如确要请显式 overwrite:true' % _dst_name)",
  'if _exist is not None:',
  '    _mark_changed()',
  '    _exist[1].RefreshTable()',
  '    _dest = _exist[0]',
  '    _reused = True',
  'else:',
  "    if _n_dst_pt > 0 and not _ctl().get('overwrite'):",
  "        raise Exception('[OVERWRITE_NOT_CONFIRMED] 目标工作表已有 %d 张透视表（%s），要再建一张请显式 overwrite:true' % (_n_dst_pt, _dst_name))",
  '    _mark_changed()',
  '    if _ws_dst is None:',
  '        _ws_dst = wb.Worksheets.Add()',
  '        _ws_dst.Name = _dst_name',
  "    _pc = wb.PivotCaches().Create(SourceType=1, SourceData=src)", // SourceType=1 xlDatabase
  "    _pt = _pc.CreatePivotTable(TableDestination=_ws_dst.Range('A1'), TableName=_tname)",
  "    for f in (args.get('rows') or []):",
  '        _pt.PivotFields(f).Orientation = 1', // xlRowField
  "    for f in (args.get('columns') or []):",
  '        _pt.PivotFields(f).Orientation = 2', // xlColumnField
  "    for f in (args.get('values') or []):",
  "        _pt.AddDataField(_pt.PivotFields(f), f + '_求和', -4157)", // xlSum
  '    _dest = _ws_dst',
  '    _reused = False',
  "output = json.dumps({'pivot': 'ok', 'dst_sheet': _dest.Name, 'table_name': _tname, 'idempotent_reuse': _reused, 'rows': args.get('rows'), 'columns': args.get('columns'), 'values': args.get('values')}, ensure_ascii=False)",
  '_finish(True)',
])

// excel_pivot_refresh — 刷新工作簿内所有透视表
const CODE_PIVOT_REFRESH = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  'wb, _opened = _resolve_wb(excel)',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  '# 先枚举一次，preview 与真实路径共用同一份目标清单，不可能对不上',
  '_targets = []',
  'for _ws in wb.Worksheets:',
  '    for _pt in _ws.PivotTables():',
  '        _targets.append((_ws.Name, _pt))',
  '_ext = 0',
  'try: _ext += wb.Connections.Count',
  'except Exception: pass',
  'for _ws in wb.Worksheets:',
  '    try: _ext += _ws.QueryTables.Count',
  '    except Exception: pass',
  "_emit_preview(_plan({'pivots': [{'sheet': s, 'name': p.Name} for s, p in _targets], 'external_sources': _ext, 'would_refresh_external': bool(_ctl().get('refresh_external_data'))}))",
  "# 外部数据连接：默认**不刷**——它会走网络，属工单 §四 点名的高风险参数，要刷得显式 refresh_external_data:true",
  "_ext_done = 0",
  'if _ctl().get("refresh_external_data"):',
  '    if _ext > 0: _mark_changed()',
  '    try:',
  '        for _c in wb.Connections:',
  '            _c.Refresh(); _ext_done += 1',
  '    except Exception: pass',
  '    for _ws in wb.Worksheets:',
  '        try:',
  '            for _qt in _ws.QueryTables:',
  '                _qt.Refresh(); _ext_done += 1',
  '        except Exception: pass',
  'count = 0',
  'for _s, _p in _targets:',
  '    if count == 0: _mark_changed()',
  '    _p.RefreshTable()',
  '    count += 1',
  "output = json.dumps({'refreshed': count, 'external_sources': _ext, 'external_refreshed': _ext_done}, ensure_ascii=False)",
  '_finish(True)',
])

// excel_journal_post — 写会计分录到账簿 + 借贷平衡校验（错则标红）
// 重跑安全：默认按批次签名识别"这批已经入过了"直接跳过；并且**每次显式设色**——
// 旧代码只在不平衡时设红、永不复位，而 ledger_gen 聚合时不看颜色，于是异常分录修正后
// 总账会把上一批重复计入。
const CODE_JOURNAL_POST = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  'wb, _opened = _resolve_wb(excel)',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  "ws = wb.Worksheets(args['sheet']) if args.get('sheet') else wb.ActiveSheet",
  'last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row', // xlUp 找 A 列最后一行
  "_has_hdr = not (last == 1 and str(ws.Cells(1,1).Value or '').strip() == '')",
  'start = (last + 1) if _has_hdr else 2',
  "entries = args.get('entries') or []",
  '_n = len(entries)',
  // 签名归一化：Cells().Value 对写进去的 '' 返回 None；金额写 int 读回 float → 比数值不比字符串。
  '# 批次签名：入账时写进去的形状，与表上末尾 N 行逐条对账',
  'def _sig_from_sheet(r):',
  "    return (_norm_date(ws.Cells(r,1).Value), str(ws.Cells(r,2).Value or '').strip(), str(ws.Cells(r,3).Value or '').strip(), float(ws.Cells(r,4).Value or 0), float(ws.Cells(r,5).Value or 0))",
  'def _sig_from_arg(e):',
  "    return (_norm_date(e.get('date','')), str(e.get('desc','') or '').strip(), str(e.get('account','') or '').strip(), float(e.get('debit') or 0), float(e.get('credit') or 0))",
  'def _same(a, b):',
  '    return a[0] == b[0] and a[1] == b[1] and a[2] == b[2] and abs(a[3]-b[3]) < 1e-9 and abs(a[4]-b[4]) < 1e-9',
  'dr = 0.0; cr = 0.0',
  'for e in entries:',
  '    _s = _sig_from_arg(e)',
  '    dr += _s[3]; cr += _s[4]',
  'balanced = abs(dr - cr) < 1e-9',
  '_want = [_sig_from_arg(e) for e in entries]',
  '_tail = [_sig_from_sheet(r) for r in range(max(2, last - _n + 1), last + 1)] if (_has_hdr and _n > 0) else []',
  '_dup = len(_tail) == _n and all(_same(_tail[i], _want[i]) for i in range(_n))',
  '_overlap = 0',
  'if not _dup and _n > 0 and _tail:',
  '    _k = min(len(_tail), _n - 1)',
  '    while _k > 0:',
  '        if all(_same(_tail[len(_tail)-_k+i], _want[i]) for i in range(_k)): break',
  '        _k -= 1',
  '    _overlap = _k',
  "_on_dup = args.get('on_duplicate') or 'skip'",
  "_emit_preview(_plan({'target': {'sheet': ws.Name, 'start_row': start, 'range': ('A%d:E%d' % (start, start + _n - 1)) if _n else None}, 'header_will_be_created': (not _has_hdr), 'entries': _n, 'debit_total': dr, 'credit_total': cr, 'balanced': balanced, 'would_mark_red': bool(_n > 0 and not balanced), 'duplicate_of_trailing': _dup, 'trailing_overlap_rows': _overlap, 'on_duplicate': _on_dup}))",
  "if _dup and _on_dup == 'skip':",
  '    # 幂等命中：不改任何东西，信封靠 _observed 如实报 changed=false',
  "    output = json.dumps({'posted': 0, 'debit_total': dr, 'credit_total': cr, 'balanced': balanced, 'idempotent_skip': True, 'duplicate_of_trailing': True}, ensure_ascii=False)",
  'else:',
  "    if _on_dup == 'replace' and _n > 0 and _has_hdr and last >= _n + 1:",
  '        # replace 是**位置覆盖**：把这批写到末尾这批的位置上。改正场景里新批次与旧批次',
  '        # 内容本就不同（摘要/金额改了），所以不能靠"批次相同"触发，得由调用方显式表态。',
  '        # 顺带复位那几行的颜色——这是唯一能清掉上一次红标的机制。',
  '        start = last - _n + 1',
  '    _mark_changed()',
  '    if not _has_hdr:',
  "        ws.Cells(1,1).Value = '日期'",
  "        ws.Cells(1,2).Value = '摘要'",
  "        ws.Cells(1,3).Value = '科目'",
  "        ws.Cells(1,4).Value = '借方'",
  "        ws.Cells(1,5).Value = '贷方'",
  '    for i, e in enumerate(entries):',
  '        r = start + i',
  "        ws.Cells(r,1).Value = e.get('date','')",
  "        ws.Cells(r,2).Value = e.get('desc','')",
  "        ws.Cells(r,3).Value = e.get('account','')",
  '        _e = _sig_from_arg(e)',
  '        ws.Cells(r,4).Value = _e[3]',
  '        ws.Cells(r,5).Value = _e[4]',
  '    if _n > 0:',
  '        _rng = ws.Range(ws.Cells(start,1), ws.Cells(start+_n-1,5))',
  '        # 每次显式设色：不平衡标红、平衡复位为自动色。不复位的话上次的红会留到这次，',
  '        # 而 ledger_gen 聚合时不看颜色 → 修正后的总账重复计入。',
  '        if not balanced:',
  '            _rng.Font.Color = 255',
  '        else:',
  '            _rng.Font.ColorIndex = -4105', // xlAutomatic
  "    output = json.dumps({'posted': _n, 'debit_total': dr, 'credit_total': cr, 'balanced': balanced, 'idempotent_skip': False, 'duplicate_of_trailing': _dup}, ensure_ascii=False)",
  '_finish(True)',
])

// excel_ledger_gen — 日记账 → 科目总账（按科目聚合借/贷，余额=借-贷公式）
// 重跑安全：目标表取确定性默认名并**按表头判归属**——自有产物原地刷新（永不索要 overwrite，
// 否则第二次相同运行也要额外参数，正好违背工单 §三 的可重复性）；用户数据才要 overwrite。
// 清理按 A:D 列既有 used range 上界，不是按新行数——否则 5 科目跑成 3 科目会留下陈旧行。
const CODE_LEDGER_GEN = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  'wb, _opened = _resolve_wb(excel)',
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
  '_order = sorted(accounts.keys())',
  '_hdr = ["科目", "借方合计", "贷方合计", "余额"]',
  "_sheet_name = args.get('output_sheet') or '总账'",
  'try: ws_out = wb.Worksheets(_sheet_name)',
  'except Exception: ws_out = None',
  '_old_last = 0',
  '_owned = False',
  '_has_content = False',
  'if ws_out is not None:',
  '    try: _old_last = ws_out.Cells(ws_out.Rows.Count, 1).End(-4162).Row',
  '    except Exception: _old_last = 0',
  '    try: _cur_hdr = [str(v or "").strip() for v in _flat(ws_out.Range("A1:D1").Value)]',
  '    except Exception: _cur_hdr = []',
  '    _owned = (_cur_hdr == _hdr)',
  '    if _old_last >= 1:',
  '        try: _has_content = any(v not in (None, "") for v in _flat(ws_out.Range("A1:D%d" % _old_last).Value))',
  '        except Exception: _has_content = True',
  '_would_clobber = bool(ws_out is not None and _has_content and not _owned)',
  '_rows = [{"account": a, "debit": accounts[a][0], "credit": accounts[a][1], "balance_formula": "=B%d-C%d" % (i+2, i+2)} for i, a in enumerate(_order)]',
  "_emit_preview(_plan({'output_sheet': _sheet_name, 'output_sheet_existed': ws_out is not None, 'tool_owned': _owned, 'entries_scanned': last - 1, 'accounts': len(_order), 'rows': _rows, 'existing_content_rows': (_old_last if _has_content else 0), 'would_clobber': _would_clobber}))",
  'if _would_clobber and not _ctl().get("overwrite"):',
  "    raise Exception('[OVERWRITE_NOT_CONFIRMED] 工作表 %s 已有内容且不是本工具生成的总账，要覆盖请显式 overwrite:true' % _sheet_name)",
  '# 逐格比对：重算结果与表上现存一致 → 什么都不用做（重跑的幂等信号）',
  '_skip = False',
  'if _owned:',
  '    _cur = _flat(ws_out.Range(ws_out.Cells(2,1), ws_out.Cells(_old_last, 3)).Value) if _old_last >= 2 else []',
  '    _new = []',
  '    for a in _order: _new.extend([a, accounts[a][0], accounts[a][1]])',
  '    if len(_cur) == len(_new):',
  '        _skip = True',
  '        for i in range(len(_new)):',
  '            if i % 3 == 0:',
  '                if str(_cur[i] or "").strip() != str(_new[i] or "").strip(): _skip = False; break',
  '            elif abs(float(_cur[i] or 0) - float(_new[i] or 0)) > 1e-9: _skip = False; break',
  'if _skip:',
  "    output = json.dumps({'accounts': len(_order), 'output_sheet': ws_out.Name, 'entries_scanned': last - 1, 'idempotent_skip': True}, ensure_ascii=False)",
  'else:',
  '    _mark_changed()',
  '    if ws_out is None:',
  '        ws_out = wb.Worksheets.Add()',
  '        ws_out.Name = _sheet_name',
  '    if _old_last >= 1:',
  '        # 只清 A:D（绝不碰 E 列及以后：用户可能在那写了备注），且清到既有 used range 上界',
  '        try: ws_out.Range(ws_out.Cells(1,1), ws_out.Cells(_old_last, 4)).ClearContents()',
  '        except Exception: pass',
  "    ws_out.Cells(1,1).Value = '科目'; ws_out.Cells(1,2).Value = '借方合计'; ws_out.Cells(1,3).Value = '贷方合计'; ws_out.Cells(1,4).Value = '余额'",
  '    r = 2',
  '    for acc in _order:',
  '        d, c = accounts[acc]',
  '        ws_out.Cells(r,1).Value = acc',
  '        ws_out.Cells(r,2).Value = d',
  '        ws_out.Cells(r,3).Value = c',
  "        ws_out.Cells(r,4).Formula = '=B%d-C%d' % (r, r)",
  '        r += 1',
  "    output = json.dumps({'accounts': len(_order), 'output_sheet': ws_out.Name, 'entries_scanned': last - 1, 'idempotent_skip': False}, ensure_ascii=False)",
  '_finish(True)',
])

// ── 插件入口 ─────────────────────────────────────────────────

// 导出内部函数供冒烟测试（DSH 只用 name/inject/apply）
export { findPython, ensureServer, runPython, finalize, classifyError, cleanupOwnedApps, runningOfficeImages, TOOL_META, resolveCtl, argsData, parseErrMeta }

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
    execute: (args, ctl) => executeRunCode(CODE_OFFICE_APPS, argsData({}, ctl)),
  })

  reg(ctx, {
    name: 'excel_formula_set',
    description: '在真实 Excel 实例中向单元格/区域写入公式（原样写入 =SUM() 等，交给 Excel 引擎计算）。path 给文件路径则先打开，否则用活动工作簿。mode=preview 时只回报目标地址、现有公式与拟写入内容，不落笔。',
    parameters: {
      path: { type: 'string', required: false, description: '工作簿完整路径（可选，缺省用活动工作簿）' },
      sheet: { type: 'string', required: false, description: '工作表名（可选，缺省用活动表）' },
      range: { type: 'string', required: true, description: '目标区域，如 A1 或 A1:B10' },
      formula: { type: 'string', required: true, description: '公式内容，如 =SUM(A1:A10)' },
      ...MODE_PARAMS,
    },
    execute: (args, ctl) => executeRunCode(CODE_FORMULA_SET, argsData(args, ctl)),
  })

  reg(ctx, {
    name: 'excel_recalc',
    description: '强制 Excel 重算全部公式并取计算后的活值（COM 是活公式链，区别于 openpyxl 写死数值的死穴）。',
    parameters: {
      path: { type: 'string', required: false, description: '工作簿完整路径（可选）' },
      sheet: { type: 'string', required: false, description: '工作表名（可选）' },
      range: { type: 'string', required: false, description: '取值的区域，如 A1（可选，缺省只重算不取值）' },
    },
    execute: (args, ctl) => executeRunCode(CODE_RECALC, argsData(args, ctl)),
  })

  reg(ctx, {
    name: 'excel_vba_run',
    description: '运行工作簿中已有的 VBA 宏（Application.Run）。高风险：宏可以改任意内容，必须显式 confirm:true（对应产品工单里的 run_vba 确认位）。mode=preview 只回报将要运行的宏与可见模块名，不执行——且宏的真实效果无法预演。',
    parameters: {
      path: { type: 'string', required: false, description: '工作簿完整路径（可选，缺省用活动工作簿）' },
      macro: { type: 'string', required: true, description: '宏名，如 Module1.MyMacro' },
      ...MODE_PARAMS,
    },
    execute: (args, ctl) => executeRunCode(CODE_VBA_RUN, argsData(args, ctl)),
  })

  reg(ctx, {
    name: 'office_launch',
    description: '显式启动本机的 Microsoft Office 应用（Excel/Word/Outlook）。',
    parameters: {
      app: { type: 'string', required: false, description: '应用名：Excel / Word / Outlook（缺省 Excel）' },
      visible: { type: 'boolean', required: false, description: '窗口是否可见（缺省 true）' },
    },
    execute: (args, ctl) => executeRunCode(CODE_LAUNCH, argsData(args, ctl)),
  })

  reg(ctx, {
    name: 'excel_new',
    description: '新建一个空 Excel 工作簿，返回工作表清单（作为后续写值/公式的目标）。',
    parameters: {},
    execute: (args, ctl) => executeRunCode(CODE_NEW, argsData({}, ctl)),
  })

  reg(ctx, {
    name: 'excel_open',
    description: '打开已有 Excel 工作簿，返回工作表清单与活动表名。',
    parameters: {
      path: { type: 'string', required: true, description: '工作簿完整路径，如 C:/path/book.xlsx' },
    },
    execute: (args, ctl) => executeRunCode(CODE_OPEN, argsData(args, ctl)),
  })

  reg(ctx, {
    name: 'excel_read_range',
    description: '读取 Excel 区域当前值（公式返回计算后活值），返回 2D 数组。',
    parameters: {
      path: { type: 'string', required: false, description: '工作簿完整路径（可选，缺省活动工作簿）' },
      sheet: { type: 'string', required: false, description: '工作表名（可选，缺省活动表）' },
      range: { type: 'string', required: true, description: '目标区域，如 A1 或 A1:C10' },
    },
    execute: (args, ctl) => executeRunCode(CODE_READ_RANGE, argsData(args, ctl)),
  })

  reg(ctx, {
    name: 'excel_write_range',
    description: '向 Excel 区域写入值：标量或与区域匹配的 2D 数组（公式请用 excel_formula_set）。mode=preview 时回报目标区域形状、输入形状与会被盖掉的非空格数量，不落笔。',
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
      ...MODE_PARAMS,
    },
    execute: (args, ctl) => executeRunCode(CODE_WRITE_RANGE, argsData(args, ctl)),
  })

  reg(ctx, {
    name: 'word_open',
    description: '打开已有 Word 文档，返回段落数、字符数与开头文本预览（COM 读取真实排版结构）。',
    parameters: {
      path: { type: 'string', required: true, description: 'Word 文档完整路径，如 C:/path/doc.docx' },
    },
    execute: (args, ctl) => executeRunCode(CODE_WORD_OPEN, argsData(args, ctl)),
  })

  reg(ctx, {
    name: 'word_edit',
    description: '在 Word 文档中全文查找替换（COM Find，wdReplaceAll，只动文本、保留原文格式）。path 给文件则先打开，否则用活动文档。mode=preview 回报命中次数与首处上下文——计数循环本身不改文档，所以预演是精确的。附着的文档默认不落盘，要落盘得显式 save:true。',
    parameters: {
      path: { type: 'string', required: false, description: 'Word 文档完整路径（可选，缺省活动文档）' },
      find: { type: 'string', required: true, description: '要查找的文本' },
      replace: { type: 'string', required: false, description: '替换成的文本（缺省删除）' },
      ...MODE_PARAMS,
    },
    execute: (args, ctl) => executeRunCode(CODE_WORD_EDIT, argsData(args, ctl)),
  })

  reg(ctx, {
    name: 'excel_pivot_create',
    description: '基于源数据区域创建 Excel 透视表（真实 PivotCache/PivotTable），返回透视表所在工作表。值字段按求和汇总。重跑安全：字段会先对着表头校验，目标工作表与表名都取确定性名字，第二次跑命中同一张表就复用刷新（idempotent_reuse:true）而不是再堆一张。mode=preview 回报表头、缺失字段、目标去向与是否会覆盖。',
    parameters: {
      path: { type: 'string', required: false, description: '工作簿完整路径（可选，缺省活动工作簿）' },
      sheet: { type: 'string', required: false, description: '源数据工作表名（可选，缺省活动表）' },
      range: { type: 'string', required: true, description: '源数据区域（含表头），如 A1:C100' },
      rows: { type: 'array', items: { type: 'string' }, required: true, description: '行字段名数组，如 ["科目"]' },
      columns: { type: 'array', items: { type: 'string' }, required: false, description: '列字段名数组，如 ["月份"]' },
      values: { type: 'array', items: { type: 'string' }, required: true, description: '值字段名数组（求和），如 ["金额"]' },
      output_sheet: { type: 'string', description: '透视表放到哪张工作表（可选，缺省「透视表」；不存在则新建）' },
      ...MODE_PARAMS,
    },
    execute: (args, ctl) => executeRunCode(CODE_PIVOT_CREATE, argsData(args, ctl)),
  })

  reg(ctx, {
    name: 'excel_pivot_refresh',
    description: '刷新工作簿内所有透视表（源数据变化后取最新汇总）。外部数据连接（QueryTable/OLEDB）默认**不刷**，要刷得显式 refresh_external_data:true——它会走网络。mode=preview 回报将刷新的透视表清单与外部源数量。',
    parameters: {
      path: { type: 'string', required: false, description: '工作簿完整路径（可选，缺省活动工作簿）' },
      refresh_external_data: { type: 'boolean', description: '是否一并刷新外部数据连接（默认 false：会走网络，属高风险参数）' },
      ...MODE_PARAMS,
    },
    execute: (args, ctl) => executeRunCode(CODE_PIVOT_REFRESH, argsData(args, ctl)),
  })

  reg(ctx, {
    name: 'excel_journal_post',
    description: '把会计分录写入 Excel 账簿（日期/摘要/科目/借方/贷方，缺表头自动建），并做借贷平衡校验（借≠贷标红）。会计旗舰场景第一环。重跑安全：会拿这批分录与账簿末尾逐条对账，命中同一批则默认跳过不重复入账（idempotent_skip:true）——真有两笔一模一样的分录请显式 on_duplicate:"append"。mode=preview 回报目标行、借贷合计、是否平衡、是否与末尾批次重复。',
    parameters: {
      path: { type: 'string', required: false, description: '工作簿完整路径（可选，缺省活动工作簿）' },
      sheet: { type: 'string', required: false, description: '账簿工作表名（可选，缺省活动表）' },
      on_duplicate: { type: 'string', enum: ['skip', 'append', 'replace'], description: '这批分录与账簿末尾批次重复时怎么办：skip 跳过（默认，幂等）/ append 照常追加（真有两笔相同分录时用）/ replace 位置覆盖——把这批写到末尾这批的位置上并复位其标红，用于改正上一次入错的分录' },
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
      ...MODE_PARAMS,
    },
    execute: (args, ctl) => executeRunCode(CODE_JOURNAL_POST, argsData(args, ctl)),
  })

  reg(ctx, {
    name: 'excel_ledger_gen',
    description: '从日记账生成科目总账：按科目聚合借贷（Python 端 SUMIF 等价），输出 科目/借方合计/贷方合计/余额（公式）。重跑安全：总账固定落在默认表「总账」并原地刷新，重算结果与表上一致时直接跳过（idempotent_skip:true），不会每次新建工作表。碰到不是本工具生成的内容才会要 overwrite。mode=preview 直接回报算好的完整总账，不落笔。',
    parameters: {
      path: { type: 'string', required: false, description: '工作簿完整路径（可选）' },
      journal_sheet: { type: 'string', required: false, description: '日记账工作表名（可选，缺省活动表）' },
      output_sheet: { type: 'string', required: false, description: '总账输出工作表名（可选，缺省「总账」；不存在则新建）' },
      ...MODE_PARAMS,
    },
    execute: (args, ctl) => executeRunCode(CODE_LEDGER_GEN, argsData(args, ctl)),
  })

  console.log(`[dsh-office-com] plugin loaded${degraded ? ' (degraded)' : ''}`)
}
