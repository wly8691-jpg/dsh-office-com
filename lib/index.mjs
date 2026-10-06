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
//   OFFICE_PYTHON   覆盖 python 命令（本插件专属，优先级最高）
//   DSH_PYTHON      全局 python 约定（与其他 Python 系插件一致，如 writing-guard）
//   OFFICE_FOLDER   OfficeMCP 工作目录（默认 D:/OfficeMCP）
//   未设上述变量时：扫 workbuddy 的 python 目录取版本最高者（不写死版本号）

import { kbAnnotateImpl, kbVerifyImpl } from './kb-annotate.mjs'
import { spawn, spawnSync } from 'node:child_process'
import { existsSync, readdirSync, readFileSync, writeFileSync, mkdirSync, realpathSync, statSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import net from 'node:net'

export const name = 'dsh-office-com'
export const inject = ['tools']

// 版本号只有这一处（package.json 另有一处，register.mjs 会断言两者一致，防漂移）。
export const VERSION = '1.2.1'

const HOST = '127.0.0.1'
const OFFICE_FOLDER = process.env.OFFICE_FOLDER || 'D:/OfficeMCP'
// 状态目录（锁文件与实例基线）。**要隔离测试就用这个变量**，别动 USERPROFILE——
// 见下面 STATE_DIR 环境变量那段注释。
const STATE_DIR = process.env.DSH_OFFICE_STATE_DIR || join(homedir(), '.dsh-office-com')
const LOCK_FILE = join(STATE_DIR, 'office.lock')
const PROBE_TIMEOUT = 20000
const CALL_TIMEOUT = 90000
const HEALTH_TIMEOUT = 1500

const PROBE_CODE = 'from officemcp.OfficeMCP import RunOfficeMCP'

// ── 自举：干净 base python（三原则）──────────────────────────

/** 剥离 PYTHONPATH（外部进程透传坏 venv 会劫持 import） */
function cleanEnv() {
  const env = { ...process.env }
  delete env.PYTHONPATH
  return env
}

/** 版本号数值比较：3.14.4 > 3.13.12（字典序会把 3.9 判成大于 3.14） */
function cmpVersion(a, b) {
  const pa = String(a).match(/\d+/g) || []
  const pb = String(b).match(/\d+/g) || []
  for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
    const d = (Number(pa[i]) || 0) - (Number(pb[i]) || 0)
    if (d) return d
  }
  return 0
}

/** workbuddy 的 python：优先读它的 current 指针，扫目录只作兜底（都不写死版本号） */
function workbuddyPython() {
  const base = join(homedir(), '.workbuddy', 'binaries', 'python', 'versions')
  const exeOf = (name) => join(base, name, 'python.exe')
  // ① current 指针才是"当前版本"的权威来源
  try {
    const cur = readFileSync(join(base, 'current'), 'utf-8').trim()
    if (cur && existsSync(exeOf(cur))) return exeOf(cur)
  } catch {
    /* 没有 current 就往下扫 */
  }
  // ② 兜底扫描：只认纯版本号目录，排除 3.13.12.old.<pid> 这类升级备份
  //    （备份里同样有 officemcp，PROBE_CODE 挡不住，必须在这里排掉）
  try {
    const found = readdirSync(base)
      .filter((name) => /^\d+(\.\d+)*$/.test(name))
      .map((name) => ({ name, exe: exeOf(name) }))
      .filter((x) => existsSync(x.exe))
      .sort((a, b) => cmpVersion(b.name, a.name))
    return found.length ? found[0].exe : null
  } catch {
    return null
  }
}

function findPython() {
  // 优先级：插件专属 > 全局约定 > 自动探测。逐个用 PROBE_CODE 试，能 import officemcp 才算数。
  const candidates = [
    process.env.OFFICE_PYTHON,
    process.env.DSH_PYTHON,
    workbuddyPython(),
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

/** 在跑的 Office 进程快照 {image: [pid]}。收尾必须按 **pid** 判归属——
 *  按镜像名判的话，起始时只要有一个 Excel 在跑，之后自己创建的也永远认不出来（全被当成"用户本来开的"）。 */
function runningOfficePids() {
  const out = {}
  for (const [img] of APP_IMAGES) {
    const pids = pidsOfImage(img)
    if (pids.length) out[img] = pids
  }
  return out
}

function cleanupOwnedApps() {
  if (!ownsServer) return // 附着别人的 SSE：那些实例不是我们起的
  const rec = readLock()
  if (!rec) return
  const pre = rec.preexisting
  // 老锁记的是镜像名数组，做不了 pid 级归属。这是**升级过渡期**才会遇到的形态
  // （只有当"旧版起的服务还活着、新版附着上去"时才会读到），此时宁可不收尾也不误杀。
  // 下次自己 spawn 时会写入新格式，自愈。
  if (!pre || Array.isArray(pre)) return
  const owned = []
  for (const [img] of APP_IMAGES) {
    const before = new Set(pre[img] || [])
    for (const pid of pidsOfImage(img)) {
      if (!before.has(pid)) owned.push(`${pid}|${img}`) // 启动前不存在 → 是我们起的
    }
  }
  if (!owned.length) return
  const py = findPython()
  if (!py) return
  // 先礼后兵。两处关键：
  //  1. **可见实例一律放过**（用户在看的），这点不能变。
  //  2. 但取 `app.Visible` 抛异常时**不能顺手 continue**——那正是僵尸对象
  //     （进程活着、对象已死、每个属性都抛），而僵尸恰恰是最该回收的一类。
  //     它取不到 pid，所以只能靠上面的 pid 归属来认。原来的写法把异常吞成
  //     `pid = None → continue`，等于僵尸永不回收。
  //  3. 落地前复核 pid 背后还是同一个 exe（1.5s 内 pid 被复用的话，宁可放过也不误杀）。
  const script = [
    'import sys, time, pythoncom, win32com.client, win32process, win32api, win32con',
    'pythoncom.CoInitialize()',
    'protected = set()',
    'for name in ("Excel", "Word", "PowerPoint"):',
    '    try:',
    '        app = win32com.client.GetActiveObject(name + ".Application")',
    '    except Exception:',
    '        continue',
    '    try:',
    '        if app.Visible:',
    '            protected.add(win32process.GetWindowThreadProcessId(app.Hwnd)[1])',
    '            continue',
    '        app.DisplayAlerts = False',
    '        app.Quit()',
    '        time.sleep(1.5)',
    '    except Exception:',
    '        pass', // 僵尸：拿不到 pid，交给下面的 pid 收尾
    'for spec in sys.argv[1:]:',
    '    pid_s, image = spec.split("|")',
    '    try:',
    '        pid = int(pid_s)',
    '    except Exception:',
    '        continue',
    '    if pid in protected:',
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
    spawnSync(py, ['-c', script, ...owned], { env: cleanEnv(), windowsHide: true, stdio: 'ignore', timeout: 15000 })
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
  // 动手前用户已经开着的 Office —— 收尾时必须放过它们。**记 pid 不记镜像名**：
  // 只记名字的话，起始时有一个 Excel 在跑就把整个镜像名标成"用户的"，
  // 之后我们自己创建的实例也永远认不出来、永不回收。
  const preexisting = runningOfficePids()
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
      clientInfo: { name: 'dsh-office-com', version: VERSION },
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
    if (result?.isError) {
      const err = text || 'RunPython isError'
      // 僵尸 COM 对象下 resetBridge() 是**没用**的：重连还是同一个 SSE 服务，
      // 而那个服务把 Officer.Excel 永久缓存在自己进程里——绑的还是那条死进程。
      // 必须把服务本身换掉，下次调用才会重新 Dispatch 出新实例。
      // 只在"我们自己起的服务"上做：附着别人的服务时杀掉它是越权。
      if (ownsServer && isDeadComError(err)) {
        try { killServer() } catch { /* 收尾尽力 */ }
        resetBridge()
      }
      return { success: false, error: err }
    }
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
const PREAMBLE = `import json, pythoncom, datetime, os, csv
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
        'finished': False, 'error': None, 'notes': [], 'extra_wbs': []}

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

# ── 会计契约（底层工具与任务级工具共用一份，防两处漂移）──
# 只有契约共享、不共享整套逻辑：底层工具的批次签名/归属判定/overwrite 闸是为"操作可能已存在的
# 工作簿"而存在的，任务级工具写的是全新输出，用不上那些。共享这四样就够，多共享反而互相拖累。
JOURNAL_HDR = ['日期', '摘要', '科目', '借方', '贷方']
LEDGER_HDR = ['科目', '借方合计', '贷方合计', '余额']
MONTHLY_HDR = ['期间', '科目', '借方', '贷方']
SUMMARY_HDR = ['排名', '项目', '金额', '占比']
DEFAULT_JOURNAL_SHEET = '日记账'
DEFAULT_LEDGER_SHEET = '总账'
DEFAULT_PIVOT_SHEET = '透视表'

def _balance_formula(r):
    return '=B%d-C%d' % (r, r)

# 源数据列名：中英皆收，CSV 表头与内联对象的键都能对上
_COL_ALIAS = {
    'date':    ['date', '日期', '交易日期'],
    'desc':    ['desc', 'summary', '摘要', '描述', '摘要说明'],
    'account': ['account', '科目', '账户', '会计科目'],
    'debit':   ['debit', '借方', '借方金额'],
    'credit':  ['credit', '贷方', '贷方金额'],
}

def _pick(row, field):
    for k in _COL_ALIAS[field]:
        for rk in row:
            if str(rk).strip().lower() == k.lower(): return row[rk]
    return None

def _num(v):
    if v in (None, ''): return 0.0
    try: return float(str(v).replace(',', '').replace('，', '').strip())
    except Exception: return 0.0

def _load_entries(args, period=None):
    # source 可以是内联数组，也可以是一份 CSV 路径。两种走同一个归一化，下游不分叉。
    src = args.get('source')
    if isinstance(src, list):
        rows = src
    else:
        if not src: raise Exception('[MISSING_PARAM] 缺少 source（CSV 路径或分录数组）')
        p = str(src).replace('/', os.sep)
        if not os.path.exists(p): raise Exception('[OPEN_FAILED] 找不到源文件: %s' % p)
        with open(p, 'r', encoding='utf-8-sig', newline='') as f:
            rows = list(csv.DictReader(f))
    out = []
    for row in rows:
        if not isinstance(row, dict): continue
        rec = {
            'date': _norm_date(_pick(row, 'date')),
            'desc': str(_pick(row, 'desc') or '').strip(),
            'account': str(_pick(row, 'account') or '').strip(),
            'debit': _num(_pick(row, 'debit')),
            'credit': _num(_pick(row, 'credit')),
        }
        if not rec['account']: continue
        if period and not rec['date'].startswith(str(period)): continue
        out.append(rec)
    return out

def _aggregate(entries):
    accounts = {}
    for e in entries:
        a = e['account']
        if a not in accounts: accounts[a] = [0.0, 0.0]
        accounts[a][0] += _num(e.get('debit')); accounts[a][1] += _num(e.get('credit'))
    return accounts

def _sheet_rows(ws, header_row=1):
    # 把一张表读成 [{列名: 值}]。列名取自表头行，空列名与全空行都跳过——
    # 任务级工具要好几种表型（月度数据/明细/汇总），读表这一段共用一份。
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    if last < header_row + 1: return [], []
    try: _ncol = int(ws.UsedRange.Columns.Count)
    except Exception: _ncol = 20
    if _ncol < 1: return [], []
    hdr = [str(v or '').strip() for v in _flat(ws.Range(ws.Cells(header_row,1), ws.Cells(header_row,_ncol)).Value)]
    hdr = [h for h in hdr if h]
    if not hdr: return [], []
    rows = []
    for _r in range(header_row + 1, last + 1):
        vals = _flat(ws.Range(ws.Cells(_r,1), ws.Cells(_r,len(hdr))).Value)
        rec = {}
        for _i, _h in enumerate(hdr):
            rec[_h] = vals[_i] if _i < len(vals) else None
        if any(v not in (None, '') for v in rec.values()): rows.append(rec)
    return hdr, rows

def _find_col(rec, name):
    # 列名容错：全角/半角空格与大小写都不计较
    if not isinstance(rec, dict): return None
    want = str(name).strip().lower()
    for k in rec:
        if str(k).strip().lower() == want: return k
    return None

def _pick_col(cols, *cands):
    for c in cands:
        for k in cols:
            if str(k).strip().lower() == str(c).strip().lower(): return k
    return None

def _sheet(wb, name):
    # 按名字取工作表。名字错了要报可读的码并列出实有表名——
    # 裸 COM 异常是 (-2147352567, '发生意外。', (0, None, None, None, 0, -2147352565), None)，
    # 那是 DISP_E_EXCEPTION 包着 DISP_E_BADINDEX，看不出"工作表不存在"，排查成本极高。
    # 返回 None 表示名字为空（调用方自行回退到 ActiveSheet 或新建）。
    if not name: return None
    try:
        return wb.Worksheets(name)
    except Exception:
        try: _names = [w.Name for w in wb.Worksheets]
        except Exception: _names = []
        raise Exception("[SHEET_NOT_FOUND] 工作表 '%s' 不存在（实有：%s）" % (name, '、'.join(_names)))

def _norm_date(v):
    # Excel 会把 '2026-01-05' 这类字符串**存成真日期**，读回来是 datetime 而不是原字符串。
    # 拿它做批次签名时只留日期部分，否则"看起来一样"的两批永远对不上（幂等静默失效）。
    if isinstance(v, (datetime.datetime, datetime.date)): return v.strftime('%Y-%m-%d')
    s = str(v or '').strip()
    if 'T' in s: return s.split('T')[0]
    if ' ' in s: return s.split(' ')[0]
    return s

def _rng(ws, addr):
    # Excel 对非法区域引用抛的是通用错误码 0x800A03EC 且**无文本**
    # （-2147352567, '发生意外。', (0, None, None, None, 0, -2146827284)），
    # 靠模糊匹配分不出来，只能在这里打上稳定码，Agent 才拿得到可分支的信息。
    try:
        return ws.Range(addr)
    except Exception:
        raise Exception('[SOURCE_RANGE_INVALID] 无法解析区域引用：%s' % addr)

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
    # 额外打开的工作簿（如套模板时同时开着模板与目标）先静默关掉，主工作簿走下面的保存/关闭流程
    for _extra in _ctx['extra_wbs']:
        try: _extra.Close(False)
        except Exception: pass
    _ctx['extra_wbs'] = []
    if _ctx['wb'] is None: return
    c = _ctl()
    if c.get('preview'):
        if _ctx['opened']: _close_wb()      # preview：只关自己开的那份，永不落盘
        return
    # 落盘许可：自己开的那份按 want_save 存；attached 下只有**显式** save:true 才存
    # （默认代存会把用户在编辑的未保存内容一并落盘，那不是他要的）。
    # 文件本就在用户手里（existing_open）一律不代存 —— 同理。
    may_save = _ctx['opened'] or (c.get('save_requested') and not _ctx['existing_open'])
    # 没有改动就不落盘：省一次无谓的写盘与 mtime 变动（幂等跳过的路径会走到这）。
    # 已经自己 SaveAs 过的（任务级工具写到指定输出路径）也不再 Save——普通 Save 会把新工作簿
    # 存到默认位置，等于写出一个用户没要的文件。
    if ok and c.get('want_save') and may_save and _ctx['changed'] and not _ctx['saved']:
        # 只读打开的工作簿：DisplayAlerts=False 下 Save() 会**静默不写**而不是抛异常，
        # 于是信封会报 saved=true 而盘上什么都没变（真机上实测到过）。
        # 所以既要先判只读，也要在 Save() 之后**复核**——不抛不等于写成功。
        try: _ro = bool(_ctx['wb'].ReadOnly)
        except Exception: _ro = False
        if _ro:
            _ctx['error'] = '[SAVE_FAILED] 工作簿以只读方式打开，改动无法落盘'
        else:
            try:
                _ctx['wb'].Save()
            except Exception as e:
                _ctx['error'] = '[SAVE_FAILED] 保存失败: %s' % _clean_err(e)
            else:
                try: _ctx['saved'] = bool(_ctx['wb'].Saved)
                except Exception: _ctx['saved'] = True
                if not _ctx['saved']:
                    _ctx['error'] = '[SAVE_FAILED] Save() 未报错但工作簿仍处于未保存状态'
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

def _clean_err(e):
    # pywin32 的 COM 异常 str() 出来是**一整个元组**
    # （hresult、'发生意外。'、(来源, 描述, helpfile, helpcontext, scode)、None），
    # 对 Agent 是纯噪音，也把 COM 细节漏到了工具外面。只取人可读的那段描述。
    a = getattr(e, 'args', None)
    if isinstance(a, tuple) and len(a) >= 3 and isinstance(a[2], tuple) and len(a[2]) >= 3 and a[2][2]:
        return str(a[2][2]).strip()
    return str(e).strip()

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
  "    raise Exception('%s [META]%s' % (_clean_err(_e), json.dumps({'partial_changes': _partial()}, ensure_ascii=False)))",
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
  kb_scan_root:        { changed: false, saved: false, verified: false },
  // 写回/回滚都自带「关掉重开核账」（属性读回一致 + 内容快照一致），verified=true 有据
  kb_write_back_office: { changed: true, saved: true, verified: true },
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
  'office_generate_accounting_report', 'office_update_monthly_report',
  'office_apply_template', 'office_prepare_management_summary',
  'office_replace_document_terms',
  'kb_taxonomy_set',
  'kb_annotate',
  'kb_write_back_office',
])

/** 任务级工具自己开文件、写到 output，而 parameters 里没有 path——
 *  模式推断必须认 output，否则会被判成 attached：它打开的工作簿没人关，直接泄漏。 */
const OUTPUT_MODE_TOOLS = new Set([
  'office_generate_accounting_report', 'office_update_monthly_report',
  'office_apply_template', 'office_prepare_management_summary',
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
  // 任务级工具没有 path、只有 output。它们自己 Workbooks.Open + SaveAs，语义就是 managed。
  const ownFile = hasPath || (OUTPUT_MODE_TOOLS.has(operation) && a.output != null && a.output !== '')
  const mode = a.mode ?? (ownFile ? 'managed' : 'attached')
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

/** 僵尸 COM 对象：进程还活着，但它占着 ROT 条目而对象已死——**每个属性访问都抛 AttributeError**，
 *  裸报错长这样：`Excel.Application.Workbooks`。杀掉那个进程即可恢复。
 *  这是背靠背跑测试时观测到的竞态：上一个进程 Quit() 后还没退干净，下一个恰好 GetActiveObject 绑上去。 */
function isDeadComError(text) {
  const t = String(text || '').trim()
  // RPC_E_DISCONNECTED（-2147417848）的实机文本是「被调用的对象已与其客户端断开连接」——
  // 与「已与该对象断开」是两种措辞，漏掉前者会让僵尸实例永久占住桥（2026-10-06 实测）。
  if (/AttributeError|_oleobj_|COM object|已与该对象断开|已与其客户端断开|对象已断开连接|被调用的对象已与/i.test(t)) return true
  // 实测到的形态：报错文本**就只有一串 COM 属性路径**（`Excel.Application.Workbooks`），
  // 连 "AttributeError" 字样都没有——那是 pywin32 在死对象上取属性时抛的原话。
  // 正常报错都有散文式描述，光秃秃一条 Application 路径只可能是这一种情况。
  return /^(Excel|Word|PowerPoint)\.Application(\.[A-Za-z_]\w*)?$/.test(t)
}

/** Python 带 [CODE] 前缀上报的稳定码——COM 报错文本随语言/版本变，模糊匹配不可靠。 */
const PREFIX_CODES = {
  MODE_INVALID: false,
  COM_OBJECT_DEAD: true, // 通道里的对象死了：换掉服务重来即可，属可重试
  OVERWRITE_NOT_CONFIRMED: false,
  RISKY_OP_NOT_CONFIRMED: false,
  SAVE_FAILED: true,      // 文件被占/只读/磁盘满多为瞬时
  VBA_FAILED: false,
  PARTIAL_WRITE: false,
  SOURCE_RANGE_INVALID: false,
  SHEET_NOT_FOUND: false,
  MISSING_PARAM: false,
  NO_ACTIVE_DOCUMENT: false,
  APP_UNAVAILABLE: false,
  EMPTY_SOURCE: false,
  OPEN_FAILED: false,
  VAULT_FORBIDDEN: false,
  VBA_ACCESS_DENIED: false,
  KB_SCAN_FAILED: false,
  DUPLICATE_ID: false,
  CATEGORY_NOT_FOUND: false,
  HAS_CHILDREN: false,
  UNKNOWN_OP: false,
  NO_TAXONOMY: false,
  UNVERIFIABLE: false,
  KB_WRITE_FAILED: false,
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
  // 僵尸 COM 对象：不认这一段的话会落到 UNKNOWN——Agent 拿到的是裸 pywin32 文本、
  // 不可重试、README 错误码表里也查不到，等于没有可行动信息。
  if (isDeadComError(e)) return { code: 'COM_OBJECT_DEAD', retryable: true }
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

/** 操作摘要：给 Agent 一行不必解析 output 就能读懂的结论。
 *  计数项从常见输出键里挑，有就写、没有就跳过——刻意保持通用，不给 21 个工具各写一份。 */
function buildSummary(operation, out, changed, saved, verified, preview) {
  const o = out && typeof out === 'object' && !Array.isArray(out) ? out : {}
  const parts = [operation]
  if (preview) parts.push('预演（未改动）')
  else if (changed) parts.push(saved ? '已修改并落盘' : '已修改（未落盘）')
  else parts.push('未改动')
  const COUNT_KEYS = ['posted', 'written', 'written_rows', 'replaced_rows', 'accounts', 'entries',
    'count', 'total_replaced', 'refreshed', 'groups', 'formula_error_count', 'kept_rows', 'rows']
  for (const k of COUNT_KEYS) {
    const v = o[k]
    if (typeof v === 'number') parts.push(`${k}=${v}`)
    else if (Array.isArray(v) && v.length) parts.push(`${k}=${v.length}`)
  }
  if (typeof o.balanced === 'boolean') parts.push(o.balanced ? '借贷平衡' : '⚠ 借贷不平衡')
  if (verified) parts.push('已校验')
  if (Array.isArray(o.flags)) for (const f of o.flags.slice(0, 3)) parts.push(f)
  return parts.join(' · ')
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
      summary: buildSummary(operation, out, changed, saved, verified, c.preview),
      output: out,
    }
  }
  const { text: rawText, meta: errMeta } = parseErrMeta((r && r.error) || 'unknown error')
  const cls = classifyError(rawText)
  // Python 侧为上报稳定码会给消息加个 [CODE] 前缀。信封已经有 error_code 字段了，
  // 文本里再来一遍是噪音（渲染出来是「失败[SAVE_FAILED]: [SAVE_FAILED] 保存失败…」）。
  const text = rawText.replace(/^\s*\[[A-Z][A-Z0-9_]+\]\s*/, '') || rawText
  const partial = errMeta.partial_changes ?? (!!meta.changed && c.mode !== 'managed' && !c.preview)
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
    partial_changes: partial,
    summary: `${operation} 失败[${cls.code}]${cls.retryable ? '（可重试）' : ''}: ${text}`,
  }
}

/** 渲染给人和 Agent 看的那一层：**先给一行摘要**，再附原始载荷。
 *  摘要是给「一眼看懂发生了什么」，载荷是给「要具体数字」——两个都需要，所以都给。 */
function renderText(_args, value) {
  if (!value) return [{ type: 'text', text: 'dsh-office-com 失败' }]
  const head = value.summary
    || (value.ok === false
      ? `失败[${value.error_code}]${value.retryable ? '（可重试）' : ''}: ${value.error}`
      : value.operation)
  if (value.ok === false) {
    return [{ type: 'text', text: value.partial_changes ? `${head}（⚠ 失败前已有部分修改残留）` : head }]
  }
  const out = value.output
  const body = typeof out === 'string' ? out : JSON.stringify(out)
  return [{ type: 'text', text: `${head}\n${body}` }]
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
    summary: { type: 'string', description: '一行操作摘要（不改动 output 载荷，专供不必解析 JSON 的快速判读）' },
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

/** B4 保险丝判据（纯函数，抽出来是为了能单测 —— 它的触发条件在正常路径上不可达，
 *  DSH 三轮验收 2026-10-06 §2.2 指出：只有"内侧核账被骗过"才会走到它。
 *
 *  原理：**必然改字节的改动**（新增属性 / 改值 / 删属性）若做完之后文件的 size+mtime
 *  **一丝未动**，那在物理上不可能 —— 说明保存没落盘（不管根因是什么）。
 *  「写同值」不计入：Office 会合理地跳过重写，计入就会误报。 */
export function kbFuseTrips(action, props, before, applied, st0, st1) {
  const byteChanging = action === 'write'
    ? Object.keys(props).filter((k) => !(k in before) || String(before[k]) !== String(props[k]))
    : (applied || []).filter((a) => a.op === 'deleted').map((a) => a.name)
  const unchanged = st0.size === st1.size && st0.mtimeNs === st1.mtimeNs
  return { byteChanging, unchanged, trips: byteChanging.length > 0 && unchanged }
}

/** 消费嵌套派发的返回 —— B4 的真因（DSH 三轮验收 2026-10-06 §三；这段是 CC 在 W4 写的）。
 *
 *  ⚠️ **工具返回失败信封时，registry 不会置 `isError`。** `finalize` 是**返回**失败信封
 *  而不是抛错，于是 dsh-tools 走 `createSuccessResult`（:3564），把信封原样放进 `value`、
 *  把 `renderText` 渲染的**散文**放进 `content`。
 *
 *  旧实现拿 `JSON.parse(content[].text)` 判成败 —— 而 `renderText` 失败时只给纯散文
 *  （`… 失败[KB_WRITE_FAILED]: …`）、成功时给"摘要行 + JSON"混合文本，**两种情况 parse 都失败**；
 *  `catch` 一吞就落到 `{ok:true, …}`。⇒ **只要工具没抛异常，缝就一律当成功。**
 *  这就是"经标注写回报成功、直连同文件报失败"的根源。
 *
 *  正解：认**结构化信号** —— registry 结果里的 `value` 就是工具自己的信封。
 *  （抽到模块级是为了能单测：这条逻辑曾经是错的，且在真机上跑了三轮才被抓到。） */
export function kbUnwrapToolResult(r) {
  if (!r || typeof r !== 'object') return { failed: true, error: 'empty tool result' }
  if (r.isError) return { failed: true, error: String(r.error || 'tool error') }
  const env = (r.value && typeof r.value === 'object') ? r.value : null   // ← 信封本体（dsh-tools:3566）
  if (env && env.ok === false) return { failed: true, error: String(env.error || 'tool failed') }
  const payload = r.output ?? r
  if (payload && payload.ok === false) return { failed: true, error: String(payload.error || 'tool failed') }
  return { failed: false, env: env || payload }
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
    // B2（DSH 验收 2026-10-06）：dsh 是以 `tool.execute(exec.arguments, exec)` 调用的
    // （dsh-tools:3310），exec 带着 signal / token / callId / rootCallId / agent / parent。
    // 原包装器只收第一个参数、**把 exec 吞掉**，于是需要嵌套派发的工具（两个 seam）拿不到
    // signal —— 而 dsh-tools:3161 会读 `exec.signal.aborted`，缺了就抛
    // `TypeError: Cannot read properties of undefined (reading 'aborted')`。
    // 现把 exec 作为**第三参**透传给工具自己的 execute（既有两个参数的工具不受影响）。
    execute: async (args, exec) => {
      // 控制面只在这里解析一次，同时喂 finalize（信封口径）和 rawExecute（下发给 Python）
      const ctl = resolveCtl(operation, args)
      if (!ctl.valid) {
        return finalize(operation, meta, args, {
          ok: false,
          error: `[MODE_INVALID] 未知 mode: ${JSON.stringify(args?.mode)}（应为 preview / managed / attached）`,
        }, ctl)
      }
      return finalize(operation, meta, args, await rawExecute(args, ctl, exec), ctl)
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
  "ws = _sheet(wb, args.get('sheet')) or wb.ActiveSheet",
  'rng = _rng(ws, args["range"])',
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
  "    v = tolist(_rng(ws, args['range']).Value)",
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
  "    raise Exception('[VBA_FAILED] 宏执行失败 %s: %s' % (args['macro'], _clean_err(_ve)))",
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
  "ws = _sheet(wb, args.get('sheet')) or wb.ActiveSheet",
  "output = json.dumps({'range': args['range'], 'value': tolist(_rng(ws, args['range']).Value)}, ensure_ascii=False)",
  '_finish(True)',
])

// excel_write_range — 写值（标量或 2D 数组；list→tuple 供 COM 赋值）
const CODE_WRITE_RANGE = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  'wb, _opened = _resolve_wb(excel)',
  "if wb is None: raise Exception('无活动工作簿且未给 path')",
  "ws = _sheet(wb, args.get('sheet')) or wb.ActiveSheet",
  'rng = _rng(ws, args["range"])',
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
  '# MatchCase 默认 True：Word 在 MatchCase=False 时会把**被找到文本的大小写**套到替换文本上——',
  '# 实测：被替换文本含大写拉丁（如 DSH）时，替换文本里的拉丁字母会被强制大写，且 ALLCAPS=False，',
  '# 是**真字符被改**而非显示格式。查找替换应当字面匹配，不该静默改字（违反 Scholarship Lock）。',
  '# 确实需要模糊匹配时显式传 match_case:false。',
  '_match_case = args.get("match_case") is not False',
  'doc, _opened = _resolve_doc(word)',
  "if doc is None: raise Exception('无活动文档且未给 path')",
  '# 先计数：纯 Find 查找不动文档（Forward + wdFindStop=0 防回绕死循环）',
  '# MatchCase 必须与下面替换那遍一致，否则"数出来 N 处、实际替换 M 处"',
  'rng = doc.Content',
  'rng.Find.ClearFormatting()',
  'rng.Find.MatchCase = _match_case',
  'count = 0',
  '_first = None',
  'while rng.Find.Execute(FindText=find_text, Forward=True, Wrap=0):',
  '    count += 1',
  '    if count == 1:',
  '        try: _first = str(rng.Paragraphs(1).Range.Text)[:200]',
  '        except Exception: _first = None',
  '    rng.Collapse(0)',
  '# 计数循环本身不碰文档，所以 preview 是白送的；第二次跑同一 find 会得到 count=0 = 天然幂等',
  "_emit_preview(_plan({'doc': doc.Name, 'find': find_text, 'replace': replace_text, 'match_case': _match_case, 'count': count, 'would_change': count > 0, 'first_hit_context': _first}))",
  '# 再替换：wdReplaceAll=2，只替换文本、保留原文格式。',
  '# 注意必须用位置参数——实测 pywin32 对 Execute(FindText=..., ReplaceWith=..., Replace=...) 具名绑定不生效（返回 True 但文本不变），位置绑定才可靠',
  'if count > 0:',
  '    _mark_changed()',
  '    doc.Content.Find.ClearFormatting()',
  '    doc.Content.Find.MatchCase = _match_case',
  '    doc.Content.Find.Execute(find_text, _match_case, False, False, False, False, True, 0, False, replace_text, 2)',
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
  "ws = _sheet(wb, args.get('sheet')) or wb.ActiveSheet",
  'src = _rng(ws, args["range"])',
  // 旧代码是 `Worksheets.Add()` 之后才在 PivotFields 上炸，attached 模式会留下孤儿工作表；
  // 而且表名按现存透视表计数（PivotTable1/2/3…），重跑必然堆积。
  '# 字段校验提到任何变更之前',
  "_fields = [str(h) for h in _flat(src.Rows(1).Value) if h not in (None, '')]",
  "_want = list(args.get('rows') or []) + list(args.get('columns') or []) + list(args.get('values') or [])",
  '_missing = [f for f in _want if f not in _fields]',
  "if _missing: raise Exception('[SOURCE_RANGE_INVALID] 源数据区域缺少字段：%s（表头实有：%s）' % ('、'.join(_missing), '、'.join(_fields)))",
  '_dst_name = args.get("output_sheet") or DEFAULT_PIVOT_SHEET',
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
  "ws = _sheet(wb, args.get('sheet')) or wb.ActiveSheet",
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
  '        for _i, _h in enumerate(JOURNAL_HDR): ws.Cells(1, _i + 1).Value = _h',
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
  "ws = _sheet(wb, args.get('journal_sheet')) or wb.ActiveSheet",
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
  '_hdr = LEDGER_HDR',
  "_sheet_name = args.get('output_sheet') or DEFAULT_LEDGER_SHEET",
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
  '    for _i, _h in enumerate(LEDGER_HDR): ws_out.Cells(1, _i + 1).Value = _h',
  '    r = 2',
  '    for acc in _order:',
  '        d, c = accounts[acc]',
  '        ws_out.Cells(r,1).Value = acc',
  '        ws_out.Cells(r,2).Value = d',
  '        ws_out.Cells(r,3).Value = c',
  '        ws_out.Cells(r,4).Formula = _balance_formula(r)',
  '        r += 1',
  "    output = json.dumps({'accounts': len(_order), 'output_sheet': ws_out.Name, 'entries_scanned': last - 1, 'idempotent_skip': False}, ensure_ascii=False)",
  '_finish(True)',
])

// office_generate_accounting_report — 任务级：一份交易数据 → 完整会计报表
// 与底层工具的区别：底层工具各管一段、各自开关文件；这一条把「写分录 → 借贷平衡 → 科目总账
// → 透视表 → 强制重算 → 读回实际值校验 → 落盘」串进**一个 COM 会话**，Agent 不必自己编排六步，
// 也不会中途留下一个只写了一半的工作簿。模板永不被改动：结果 SaveAs 到 output。
const CODE_ACCOUNTING_REPORT = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  'excel.DisplayAlerts = False',
  "if not args.get('output'): raise Exception('[MISSING_PARAM] 缺少 output（报告落盘路径）')",
  "out_path = str(args['output']).replace('/', os.sep)",
  "entries = _load_entries(args, args.get('period'))",
  "if not entries: raise Exception('[EMPTY_SOURCE] 源数据为空，或该期间无记录')",
  'dr = sum(_num(e.get("debit")) for e in entries)',
  'cr = sum(_num(e.get("credit")) for e in entries)',
  'balanced = abs(dr - cr) < 1e-9',
  'accounts = _aggregate(entries)',
  '_order = sorted(accounts.keys())',
  "_jname = args.get('journal_sheet') or DEFAULT_JOURNAL_SHEET",
  "_lname = args.get('ledger_sheet') or DEFAULT_LEDGER_SHEET",
  "_pname = args.get('pivot_sheet') or DEFAULT_PIVOT_SHEET",
  '_out_exists = os.path.exists(out_path)',
  '_tpl = args.get("template")',
  '_src_desc = (args.get("source") if not isinstance(args.get("source"), list) else "<内联 %d 条>" % len(args["source"]))',
  "_emit_preview(_plan({'source': _src_desc, 'period': args.get('period'), 'entries': len(entries), 'accounts': len(_order), 'debit_total': dr, 'credit_total': cr, 'balanced': balanced, 'sheets': {'journal': _jname, 'ledger': _lname, 'pivot': _pname}, 'template': _tpl, 'output_path': out_path, 'output_exists': _out_exists, 'would_overwrite': _out_exists, 'ledger_rows': [{'account': a, 'debit': accounts[a][0], 'credit': accounts[a][1]} for a in _order]}))",
  '# 预演先出（它要报告 would_overwrite），守卫放后面',
  "if _out_exists and not _ctl().get('overwrite'):",
  "    raise Exception('[OVERWRITE_NOT_CONFIRMED] 输出文件已存在（%s），要覆盖请显式 overwrite:true' % out_path)",
  '_mark_changed()',
  'if _tpl:',
  "    tpl = str(_tpl).replace('/', os.sep)",
  "    if not os.path.exists(tpl): raise Exception('[OPEN_FAILED] 找不到模板: %s' % tpl)",
  '    wb = excel.Workbooks.Open(tpl)',
  'else:',
  '    wb = excel.Workbooks.Add()',
  "_ctx['wb'] = wb; _ctx['opened'] = True", // 自己开的，收尾关它
  'try:',
  '    ws_j = wb.Worksheets(_jname)',
  'except Exception:',
  '    ws_j = wb.Worksheets.Add(); ws_j.Name = _jname',
  'ws_j.Cells.ClearContents()',
  'for _i, _h in enumerate(JOURNAL_HDR): ws_j.Cells(1, _i + 1).Value = _h',
  'for _i, _e in enumerate(entries):',
  '    _r = 2 + _i',
  "    ws_j.Cells(_r,1).Value = _e.get('date','')",
  "    ws_j.Cells(_r,2).Value = _e.get('desc','')",
  "    ws_j.Cells(_r,3).Value = _e['account']",
  "    ws_j.Cells(_r,4).Value = _num(_e.get('debit'))",
  "    ws_j.Cells(_r,5).Value = _num(_e.get('credit'))",
  'if not balanced:',
  '    ws_j.Range(ws_j.Cells(2,1), ws_j.Cells(1 + len(entries), 5)).Font.Color = 255',
  'try:',
  '    ws_l = wb.Worksheets(_lname)',
  'except Exception:',
  '    ws_l = wb.Worksheets.Add(); ws_l.Name = _lname',
  'ws_l.Cells.ClearContents()',
  'for _i, _h in enumerate(LEDGER_HDR): ws_l.Cells(1, _i + 1).Value = _h',
  '_r = 2',
  'for _a in _order:',
  '    _d, _c = accounts[_a]',
  '    ws_l.Cells(_r,1).Value = _a',
  '    ws_l.Cells(_r,2).Value = _d',
  '    ws_l.Cells(_r,3).Value = _c',
  '    ws_l.Cells(_r,4).Formula = _balance_formula(_r)',
  '    _r += 1',
  'try:',
  '    ws_p = wb.Worksheets(_pname)',
  'except Exception:',
  '    ws_p = wb.Worksheets.Add(); ws_p.Name = _pname',
  "_src_rng = 'A1:E%d' % (len(entries) + 1)",
  "_tname = ('PT_' + _jname + '_' + _src_rng).replace(' ', '_').replace('$', '').replace(':', '_').replace('/', '_')[:200]",
  '_exist = None',
  'for _w in wb.Worksheets:',
  '    for _pt in _w.PivotTables():',
  '        if _pt.Name == _tname: _exist = _pt',
  'if _exist is not None:',
  '    _exist.RefreshTable()',
  'else:',
  '    _pc = wb.PivotCaches().Create(SourceType=1, SourceData=ws_j.Range(_src_rng))',
  "    _pt = _pc.CreatePivotTable(TableDestination=ws_p.Range('A1'), TableName=_tname)",
  "    _pt.PivotFields('科目').Orientation = 1",
  "    for _f in ('借方', '贷方'):",
  "        _pt.AddDataField(_pt.PivotFields(_f), _f + '_求和', -4157)",
  'excel.Calculate()', // 强制重算，下面读的是活值不是写死的数
  '_verify = []',
  'for _i, _a in enumerate(_order):',
  '    _r = 2 + _i',
  "    _verify.append({'account': _a, 'debit': ws_l.Cells(_r,2).Value, 'credit': ws_l.Cells(_r,3).Value, 'balance': ws_l.Cells(_r,4).Value})",
  'wb.SaveAs(out_path)',
  "_ctx['saved'] = True", // 已自行落盘，别让 _finish 再 Save 一次（那会存到默认位置）
  "output = json.dumps({'report': True, 'output_path': out_path, 'journal': {'sheet': _jname, 'posted': len(entries), 'debit_total': dr, 'credit_total': cr, 'balanced': balanced}, 'ledger': {'sheet': _lname, 'accounts': len(_order)}, 'pivot': {'sheet': ws_p.Name, 'table': _tname}, 'verification': _verify, 'recalculated': True}, ensure_ascii=False)",
  '_finish(True)',
])

// office_check_workbook — 任务级：动手前的体检（只读，不改任何东西）
// 给 Agent 一个「能不能安全地改这份工作簿」的判断：有没有公式已经是错的、有没有外部链接、
// 有没有被保护、有没有未保存改动、体量多大。xlCellTypeFormulas/xlErrors 一次调用定位所有
// 求值出错的公式单元格，比逐格判断快几个数量级。
const CODE_CHECK_WORKBOOK = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  'excel.DisplayAlerts = False',
  "if not args.get('path'): raise Exception('[MISSING_PARAM] 缺少 path')",
  'wb, _opened = _resolve_wb(excel)',
  "if wb is None: raise Exception('[OPEN_FAILED] 无法打开工作簿')",
  'excel.Calculate()', // 先重算，否则读到的是上次存盘时的缓存值，错误可能已经不存在
  '_sheets = []',
  '_total_err = 0',
  '_errs = []',
  '_want = args.get("sheet")',
  'for _ws in wb.Worksheets:',
  '    if _want and _ws.Name != _want: continue',
  '    try: _used = _addr(_ws.UsedRange)',
  '    except Exception: _used = None',
  '    # -4123 = xlCellTypeFormulas, 16 = xlErrors：一次拿到所有求值出错的公式单元格。',
  '    # 注意返回的是**多区域 Range**，相邻的错格会连成一片（D2/E2 → $D$2:$E$2），',
  '    # 直接把 Address 当单元格用会把两个错算成一个。所以按 Area 拆开，小区域逐格展开。',
  '    try: _hit = _ws.UsedRange.SpecialCells(-4123, 16)',
  '    except Exception: _hit = None',
  '    _bad = []',
  '    if _hit is not None:',
  '        try: _areas = list(_hit.Areas)',
  '        except Exception: _areas = [_hit]',
  '        for _a in _areas:',
  '            try: _n = int(_a.Cells.Count)',
  '            except Exception: _n = 1',
  '            if _n <= 20:',
  '                try:',
  '                    for _c in _a.Cells: _bad.append(_addr(_c))',
  '                except Exception: _bad.append(_addr(_a))',
  '            else:',
  "                _bad.append('%s（%d 格）' % (_addr(_a), _n))",
  '    for _c in _bad: _errs.append({"sheet": _ws.Name, "cell": _c})',
  '    _total_err += len(_bad)',
  '    try: _ptn = len(list(_ws.PivotTables()))',
  '    except Exception: _ptn = 0',
  '    try: _qtn = _ws.QueryTables.Count',
  '    except Exception: _qtn = 0',
  '    _sheets.append({"name": _ws.Name, "used_range": _used, "pivot_tables": _ptn, "query_tables": _qtn, "protected": bool(_ws.ProtectContents), "visible": bool(_ws.Visible), "formula_errors": len(_bad)})',
  'try: _links = [str(x) for x in (wb.LinkSources(1) or [])]',
  'except Exception: _links = []',
  '_names = []',
  'try:',
  '    for _n in wb.Names: _names.append(str(_n.Name))',
  'except Exception: pass',
  'try: _dirty = (not bool(wb.Saved))',
  'except Exception: _dirty = None',
  'try: _ro = bool(wb.ReadOnly)',
  'except Exception: _ro = False',
  '_edit_ok = bool((not _ro) and (not _links) and _total_err == 0 and not any(s["protected"] for s in _sheets))',
  "_flags = []",
  "if _ro: _flags.append('工作簿以只读方式打开，改动无法落盘')",
  "if _dirty: _flags.append('有未保存改动')",
  "if _links: _flags.append('含 %d 个外部链接，改动前需确认链接可达' % len(_links))",
  "if _total_err: _flags.append('有 %d 个公式求值出错' % _total_err)",
  "if any(s['protected'] for s in _sheets): _flags.append('有工作表被保护')",
  "if any(s['query_tables'] for s in _sheets): _flags.append('有外部数据查询表（刷新会走网络）')",
  "output = json.dumps({'workbook': wb.Name, 'path': wb.FullName, 'sheets': _sheets, 'formula_errors': _errs[:50], 'formula_error_count': _total_err, 'external_links': _links, 'defined_names': _names[:50], 'dirty': _dirty, 'read_only': _ro, 'safe_to_edit': _edit_ok, 'flags': _flags}, ensure_ascii=False)",
  '_finish(True)',
])

// office_replace_document_terms — 任务级：Word 文档多组术语批量替换
// 与 word_edit 的区别：一次给一组术语（合同里的甲方/乙方/金额/日期…），逐条报命中数，
// 且**先全部计数、再统一替换**——先替换会让后面的术语对着已改过的文本计数，数字对不上。
const CODE_REPLACE_TERMS = comCode([
  'word = Officer.Word',
  "if not word: raise Exception('Word 未安装/不可用')",
  'word.DisplayAlerts = 0',
  "terms = args.get('terms')",
  '_pairs = []',
  'if isinstance(terms, dict):',
  '    for _k, _v in terms.items(): _pairs.append((str(_k), str(_v or "")))',
  'elif isinstance(terms, list):',
  '    for _t in terms:',
  "        if isinstance(_t, dict) and _t.get('find'): _pairs.append((str(_t['find']), str(_t.get('replace') or '')))",
  "if not _pairs: raise Exception('[MISSING_PARAM] 缺少 terms（格式：[{find,replace}] 或 {find: replace}）')",
  'doc, _opened = _resolve_doc(word)',
  "if doc is None: raise Exception('[NO_ACTIVE_DOCUMENT] 无活动文档且未给 path')",
  '# MatchCase 默认 True —— 同 word_edit：MatchCase=False 时 Word 会把被找到文本的大小写套到',
  '# 替换文本上，静默改字。两遍必须同口径，否则"数出 N 处、实际替换 M 处"。',
  '_match_case = args.get("match_case") is not False',
  '# 第一遍：全部只计数，不动文档（纯 Find，Forward + Wrap=0 防回绕死循环）',
  '_counts = []',
  'for _f, _r in _pairs:',
  '    _rng = doc.Content',
  '    _rng.Find.ClearFormatting()',
  '    _rng.Find.MatchCase = _match_case',
  '    _c = 0',
  '    _first = None',
  '    while _rng.Find.Execute(FindText=_f, Forward=True, Wrap=0):',
  '        _c += 1',
  '        if _c == 1:',
  '            try: _first = str(_rng.Paragraphs(1).Range.Text)[:200]',
  '            except Exception: _first = None',
  '        _rng.Collapse(0)',
  "    _counts.append({'find': _f, 'replace': _r, 'count': _c, 'first_hit_context': _first})",
  "_total = sum(x['count'] for x in _counts)",
  "_emit_preview(_plan({'doc': doc.Name, 'terms': _counts, 'total_hits': _total, 'would_change': _total > 0, 'note': '计数是替换前统计的；术语之间有重叠时实际替换数可能不同'}))",
  '# 第二遍：统一替换（wdReplaceAll=2，只动文本、保留原文格式）。',
  '# 必须用位置参数——实测 pywin32 对 Execute(FindText=..., ReplaceWith=..., Replace=...) 具名绑定不生效（返回 True 但文本不变）',
  'for _item in _counts:',
  "    if _item['count'] > 0:",
  '        _mark_changed()',
  '        doc.Content.Find.ClearFormatting()',
  "        doc.Content.Find.MatchCase = _match_case",
  "        doc.Content.Find.Execute(_item['find'], _match_case, False, False, False, False, True, 0, False, _item['replace'], 2)",
  "output = json.dumps({'doc': doc.Name, 'terms': [{'find': x['find'], 'replace': x['replace'], 'count': x['count']} for x in _counts], 'total_replaced': _total}, ensure_ascii=False)",
  '_finish(True)',
])

// office_update_monthly_report — 任务级：按期间更新月度报告里的一个期间
// 幂等口径：同一个期间重跑是**替换那一期间的行**，不是追加；其他期间一行不动。
// 原报告只读，结果写到 output —— 避免"更新报告"变成一次不可逆的原地覆写。
const CODE_UPDATE_MONTHLY = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  'excel.DisplayAlerts = False',
  "if not args.get('report'): raise Exception('[MISSING_PARAM] 缺少 report（要更新的报告路径）')",
  "if not args.get('output'): raise Exception('[MISSING_PARAM] 缺少 output（结果落盘路径）')",
  "period = str(args.get('period') or '').strip()",
  "if not period: raise Exception('[MISSING_PARAM] 缺少 period（YYYY-MM）')",
  "rpt = str(args['report']).replace('/', os.sep)",
  "out_path = str(args['output']).replace('/', os.sep)",
  "if not os.path.exists(rpt): raise Exception('[OPEN_FAILED] 找不到报告: %s' % rpt)",
  "entries = _load_entries(args, period)",
  "if not entries: raise Exception('[EMPTY_SOURCE] 源数据为空，或该期间无记录')",
  "_sname = args.get('sheet') or '月度数据'",
  "_out_exists = os.path.exists(out_path)",
  '_wb = excel.Workbooks.Open(rpt)',
  "_ctx['wb'] = _wb; _ctx['opened'] = True",
  'try:',
  '    ws = _wb.Worksheets(_sname)',
  'except Exception:',
  '    ws = _wb.Worksheets.Add(); ws.Name = _sname',
  '_hdr, _rows = _sheet_rows(ws)',
  '_c_period = _pick_col(_hdr, "期间", "月份", "period") if _hdr else None',
  '_old = [r for r in _rows if _c_period and str(r.get(_c_period) or "").strip() == period]',
  '_kept = [r for r in _rows if not (_c_period and str(r.get(_c_period) or "").strip() == period)]',
  "_c_acc = _pick_col(_hdr, '科目', 'account') if _hdr else None",
  "_c_dr = _pick_col(_hdr, '借方', 'debit') if _hdr else None",
  "_c_cr = _pick_col(_hdr, '贷方', 'credit') if _hdr else None",
  '_dr = sum(_num(e.get("debit")) for e in entries)',
  '_cr = sum(_num(e.get("credit")) for e in entries)',
  "_emit_preview(_plan({'report': rpt, 'sheet': _sname, 'period': period, 'existing_rows_for_period': len(_old), 'other_period_rows_kept': len(_kept), 'incoming_entries': len(entries), 'debit_total': _dr, 'credit_total': _cr, 'balanced': abs(_dr-_cr) < 1e-9, 'output_path': out_path, 'output_exists': _out_exists, 'is_first_time_period': len(_old) == 0}))",
  "if _out_exists and not _ctl().get('overwrite'):",
  "    raise Exception('[OVERWRITE_NOT_CONFIRMED] 输出文件已存在（%s），要覆盖请显式 overwrite:true' % out_path)",
  '# 归属判定：表非空但表头不像月度表就不是我们的产物，一律不碰。',
  '# 不拦的话，下面那句 Cells.ClearContents 会把用户自己的报表直接清空。',
  '# 判据要覆盖**全部**必需列，不能只看「期间」：只有「期间」而缺「借方/贷方」时最隐蔽——',
  '# 清空后金额列会被按 MONTHLY_HDR 写成 null（_row.get("借方") 取不到），却仍报 kept_rows 成功，静默毁数据。',
  '_missing_hdr = []',
  'if _rows:',
  '    if _c_period is None: _missing_hdr.append("期间")',
  '    if _c_dr is None: _missing_hdr.append("借方")',
  '    if _c_cr is None: _missing_hdr.append("贷方")',
  '_foreign = bool(_rows) and bool(_missing_hdr)',
  "if _foreign and not _ctl().get('overwrite'):",
  "    raise Exception('[SCHEMA_MISMATCH] 工作表 %s 已有内容但不是月度表：缺列 %s（本表要求 %s），实有列 %s。请改用规范表头，或显式 overwrite:true 覆盖' % (_sname, '、'.join(_missing_hdr), '、'.join(MONTHLY_HDR), '、'.join(_hdr)))",
  '_mark_changed()',
  'ws.Cells.ClearContents()',
  'for _i, _h in enumerate(MONTHLY_HDR): ws.Cells(1, _i + 1).Value = _h',
  '_r = 2',
  'for _row in _kept:',
  '    for _i, _h in enumerate(MONTHLY_HDR):',
  '        ws.Cells(_r, _i + 1).Value = _row.get(_h) if isinstance(_row, dict) else None',
  '    _r += 1',
  'for _e in entries:',
  "    ws.Cells(_r,1).Value = period",
  "    ws.Cells(_r,2).Value = _e['account']",
  "    ws.Cells(_r,3).Value = _num(_e.get('debit'))",
  "    ws.Cells(_r,4).Value = _num(_e.get('credit'))",
  '    _r += 1',
  'excel.Calculate()',
  'try: _ws_total = _wb.Worksheets(args.get("total_sheet") or "总览")',
  'except Exception: _ws_total = None',
  '_total_rows = None',
  'if _ws_total is not None:',
  '    _h2, _rows2 = _sheet_rows(_ws_total)',
  '    _total_rows = len(_rows2)',
  '_wb.SaveAs(out_path)',
  "_ctx['saved'] = True",
  "output = json.dumps({'report': True, 'output_path': out_path, 'sheet': _sname, 'period': period, 'replaced_rows': len(_old), 'kept_rows': len(_kept), 'written_rows': len(entries), 'debit_total': _dr, 'credit_total': _cr, 'balanced': abs(_dr-_cr) < 1e-9, 'is_first_time_period': len(_old) == 0, 'total_sheet_rows': _total_rows}, ensure_ascii=False)",
  '_finish(True)',
])

// office_apply_template — 任务级：只搬格式，不搬数据
// 把模板工作表的**外观**（字体/颜色/边框/数字格式，以及列宽行高与冻结窗格）套到目标工作簿的
// 对应表上；目标表的数据一行不动。数据与外观分家，才不会出现"套模板把数据弄丢"。
const CODE_APPLY_TEMPLATE = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  'excel.DisplayAlerts = False',
  "if not args.get('template'): raise Exception('[MISSING_PARAM] 缺少 template')",
  "if not args.get('target'): raise Exception('[MISSING_PARAM] 缺少 target（要套格式的工作簿）')",
  "if not args.get('output'): raise Exception('[MISSING_PARAM] 缺少 output（结果落盘路径）')",
  "tpl = str(args['template']).replace('/', os.sep)",
  "tgt = str(args['target']).replace('/', os.sep)",
  "out_path = str(args['output']).replace('/', os.sep)",
  "for _p in (tpl, tgt):",
  "    if not os.path.exists(_p): raise Exception('[OPEN_FAILED] 找不到文件: %s' % _p)",
  "_map = args.get('sheet_map') or {}",
  "_out_exists = os.path.exists(out_path)",
  '_wb_t = excel.Workbooks.Open(tpl)',
  '_wb_g = excel.Workbooks.Open(tgt)',
  "_ctx['wb'] = _wb_g; _ctx['opened'] = True", // 目标走正常收尾
  "_ctx['extra_wbs'].append(_wb_t)", // 模板也由我们开，收尾一并关掉（异常路径也不漏）
  '_tnames = [w.Name for w in _wb_t.Worksheets]',
  '_gnames = [w.Name for w in _wb_g.Worksheets]',
  '_plan_pairs = []',
  'for _tn in _tnames:',
  '    _gn = _map.get(_tn, _tn)',
  "    _plan_pairs.append({'template_sheet': _tn, 'target_sheet': _gn, 'target_exists': _gn in _gnames})",
  "_emit_preview(_plan({'template': tpl, 'target': tgt, 'output_path': out_path, 'output_exists': _out_exists, 'pairs': _plan_pairs, 'note': '只搬格式（字体/颜色/边框/数字格式/列宽行高/冻结窗格），目标表的数据不动'}))",
  "if _out_exists and not _ctl().get('overwrite'):",
  "    raise Exception('[OVERWRITE_NOT_CONFIRMED] 输出文件已存在（%s），要覆盖请显式 overwrite:true' % out_path)",
  '_mark_changed()',
  '_applied = []',
  'for _p in _plan_pairs:',
  '    _ts = _p["template_sheet"]; _gs = _p["target_sheet"]',
  '    if not _p["target_exists"]:',
  '        _nw = _wb_g.Worksheets.Add()',
  '        _nw.Name = _gs',
  "    _src = _wb_t.Worksheets(_ts)",
  "    _dst = _wb_g.Worksheets(_gs)",
  '    _items = []',
  '    try:',
  '        _src.UsedRange.Copy()',
  "        _dst.Range('A1').PasteSpecial(-4122)", // xlPasteFormats：一次搬完字体/颜色/边框/数字格式
  "        _items.append('formats')",
  '    except Exception: pass',
  '    try:',
  '        _nc = int(_src.UsedRange.Columns.Count)',
  '        for _i in range(1, _nc + 1):',
  '            _dst.Columns(_i).ColumnWidth = _src.Columns(_i).ColumnWidth',
  "        _items.append('column_widths')",
  '    except Exception: pass',
  '    try:',
  '        _nr = int(_src.UsedRange.Rows.Count)',
  '        for _i in range(1, _nr + 1):',
  '            _dst.Rows(_i).RowHeight = _src.Rows(_i).RowHeight',
  "        _items.append('row_heights')",
  '    except Exception: pass',
  '    try:',
  "        _wb_t.Activate(); _src.Activate()",
  '        _fr = int(excel.ActiveWindow.SplitRow or 0)',
  '        _fc = int(excel.ActiveWindow.SplitColumn or 0)',
  "        _wb_g.Activate(); _dst.Activate()",
  '        excel.ActiveWindow.FreezePanes = False',
  '        if _fr: excel.ActiveWindow.SplitRow = _fr',
  '        if _fc: excel.ActiveWindow.SplitColumn = _fc',
  '        if _fr or _fc: excel.ActiveWindow.FreezePanes = True',
  "        _items.append('freeze_panes')",
  '    except Exception: pass',
  "    _applied.append({'target_sheet': _gs, 'applied': _items})",
  'excel.Calculate()',
  '_wb_g.SaveAs(out_path)',
  "_ctx['saved'] = True; _ctx['wb'] = _wb_g",
  "output = json.dumps({'applied': True, 'output_path': out_path, 'pairs': _applied, 'template_unchanged': True}, ensure_ascii=False)",
  '_finish(True)',
])

// office_prepare_management_summary — 任务级：从数据表出一张管理层摘要
// 按维度聚合 + 排行 + 占比，可选与上一期间对比。只用真实数据算，不编造结论。
const CODE_MGMT_SUMMARY = comCode([
  'excel = Officer.Excel',
  "if not excel: raise Exception('Excel 未安装/不可用')",
  'excel.DisplayAlerts = False',
  "if not args.get('source'): raise Exception('[MISSING_PARAM] 缺少 source（数据工作簿路径）')",
  "if not args.get('output'): raise Exception('[MISSING_PARAM] 缺少 output（结果落盘路径）')",
  "src_p = str(args['source']).replace('/', os.sep)",
  "out_path = str(args['output']).replace('/', os.sep)",
  "if not os.path.exists(src_p): raise Exception('[OPEN_FAILED] 找不到数据文件: %s' % src_p)",
  "_top_n = int(args.get('top_n') or 10)",
  "_sname = args.get('sheet')",
  "_out_exists = os.path.exists(out_path)",
  '_wb = excel.Workbooks.Open(src_p)',
  "_ctx['wb'] = _wb; _ctx['opened'] = True",
  "_ws = _sheet(_wb, _sname) or _wb.ActiveSheet",
  '_hdr, _rows = _sheet_rows(_ws)',
  "if not _rows: raise Exception('[EMPTY_SOURCE] 数据表没有数据行')",
  "_gc = _pick_col(_hdr, args.get('group_by') or '科目', '科目', '部门', '项目', 'account', 'group')",
  "_mc = _pick_col(_hdr, args.get('measure') or '借方', '借方', '金额', 'amount', 'debit')",
  "_pc = _pick_col(_hdr, '期间', '月份', 'period')",
  "if not _gc: raise Exception('[SOURCE_RANGE_INVALID] 找不到分组列（可用列：%s）' % '、'.join(_hdr))",
  "if not _mc: raise Exception('[SOURCE_RANGE_INVALID] 找不到度量列（可用列：%s）' % '、'.join(_hdr))",
  '_agg = {}',
  '_periods = {}',
  'for _row in _rows:',
  '    _g = str(_row.get(_gc) or "").strip()',
  '    if not _g: continue',
  '    _v = _num(_row.get(_mc))',
  '    _agg[_g] = _agg.get(_g, 0.0) + _v',
  '    if _pc:',
  '        _p = str(_row.get(_pc) or "").strip()',
  '        _periods.setdefault(_g, {})[_p] = _periods.get(_g, {}).get(_p, 0.0) + _v',
  '_ranked = sorted(_agg.items(), key=lambda kv: kv[1], reverse=True)',
  '_total = sum(_agg.values())',
  '_period_keys = sorted({k for v in _periods.values() for k in v})',
  "_emit_preview(_plan({'source': src_p, 'sheet': _ws.Name, 'group_by': _gc, 'measure': _mc, 'groups': len(_agg), 'total': _total, 'top': [{'name': k, 'value': v} for k, v in _ranked[:_top_n]], 'periods': _period_keys, 'output_path': out_path, 'output_exists': _out_exists}))",
  "if _out_exists and not _ctl().get('overwrite'):",
  "    raise Exception('[OVERWRITE_NOT_CONFIRMED] 输出文件已存在（%s），要覆盖请显式 overwrite:true' % out_path)",
  '_mark_changed()',
  "_ws2name = args.get('summary_sheet') or '管理层摘要'",
  'try:',
  '    ws_s = _wb.Worksheets(_ws2name)',
  'except Exception:',
  '    ws_s = _wb.Worksheets.Add(); ws_s.Name = _ws2name',
  'ws_s.Cells.ClearContents()',
  "ws_s.Cells(1,1).Value = '管理层摘要'",
  "ws_s.Cells(2,1).Value = '分组维度'; ws_s.Cells(2,2).Value = _gc",
  "ws_s.Cells(3,1).Value = '度量'; ws_s.Cells(3,2).Value = _mc",
  "ws_s.Cells(4,1).Value = '分组数'; ws_s.Cells(4,2).Value = len(_agg)",
  "ws_s.Cells(5,1).Value = '合计'; ws_s.Cells(5,2).Value = _total",
  'for _i, _h in enumerate(SUMMARY_HDR): ws_s.Cells(7, _i + 1).Value = _h',
  '_r = 8',
  'for _i, (_k, _v) in enumerate(_ranked[:_top_n]):',
  '    ws_s.Cells(_r,1).Value = _i + 1',
  '    ws_s.Cells(_r,2).Value = _k',
  '    ws_s.Cells(_r,3).Value = _v',
  '    ws_s.Cells(_r,4).Value = (_v / _total) if _total else 0',
  '    ws_s.Cells(_r,4).NumberFormat = "0.00%"',
  '    _r += 1',
  "_cp = args.get('compare_period')",
  '_cmp = None',
  'if _pc and _period_keys and _cp:',
  '    _idx = _period_keys.index(_cp) if _cp in _period_keys else len(_period_keys) - 1',
  '    if _idx > 0:',
  '        _cur, _prev = _period_keys[_idx], _period_keys[_idx - 1]',
  '        _cmp = []',
  '        for _k, _v in _ranked[:_top_n]:',
  '            _a = _periods.get(_k, {}).get(_prev, 0.0)',
  '            _b = _periods.get(_k, {}).get(_cur, 0.0)',
  '            _cmp.append({\'项目\': _k, \'上期\': _a, \'本期\': _b, \'变动\': _b - _a})',
  '        _r += 2',
  "        ws_s.Cells(_r,1).Value = '期间对比（%s → %s）' % (_prev, _cur)",
  '        _r += 1',
  "        for _i, _h in enumerate(['项目', '上期', '本期', '变动']): ws_s.Cells(_r, _i + 1).Value = _h",
  '        _r += 1',
  '        for _c in _cmp:',
  '            ws_s.Cells(_r,1).Value = _c[\'项目\']',
  '            ws_s.Cells(_r,2).Value = _c[\'上期\']',
  '            ws_s.Cells(_r,3).Value = _c[\'本期\']',
  '            ws_s.Cells(_r,4).Value = _c[\'变动\']',
  '            _r += 1',
  'excel.Calculate()',
  '_wb.SaveAs(out_path)',
  "_ctx['saved'] = True",
  "output = json.dumps({'summary': True, 'output_path': out_path, 'summary_sheet': ws_s.Name, 'group_by': _gc, 'measure': _mc, 'groups': len(_agg), 'total': _total, 'top': [{'rank': i + 1, 'name': k, 'value': v, 'share': (v / _total) if _total else 0} for i, (k, v) in enumerate(_ranked[:_top_n])], 'periods': _period_keys, 'comparison': _cmp}, ensure_ascii=False)",
  '_finish(True)',
])

// ── 插件入口 ─────────────────────────────────────────────────

// 导出内部函数供冒烟测试（DSH 只用 name/inject/apply）
export { findPython, ensureServer, runPython, finalize, classifyError, cleanupOwnedApps, runningOfficePids, TOOL_META, resolveCtl, argsData, parseErrMeta }

/** @param {import('@deepseek-ai/cordis').Context} ctx */
let _ctx = null

export function apply(ctx) {
  _ctx = ctx
  // 启动即探：officemcp 可用性决定是否降级（不阻塞 apply）
  const py = findPython()
  if (!py) {
    degraded = true
    degradeReason = '本机未装 OfficeMCP（officemcp）：COM 通道不可用，所有工具都会返回这个报错（无 JS 回退实现）'
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
      match_case: { type: 'boolean', required: false, description: '是否区分大小写（默认 true）。默认字面匹配：Word 在 MatchCase=false 时会把被找到文本的大小写套到替换文本上，静默改字（如被替换文本含 DSH 时，替换文本的 latin 会被强制大写）' },
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

  // ── 任务级工具（v0.4）：把几十个 COM 步骤收进一次调用 ──────────
  reg(ctx, {
    name: 'office_generate_accounting_report',
    description: '【任务级】一份交易数据 → 一整套会计报表：写入日记账 → 借贷平衡校验（借≠贷标红）→ 生成科目总账（余额活公式）→ 按科目建透视表 → 强制重算 → 读回实际值校验 → 落盘。模板永不被改动，结果写到 output。mode=preview 时不落笔，回报过滤后条数、借贷合计、目标工作表与输出是否已存在。重复生成同一输出需显式 overwrite:true。',
    parameters: {
      source: { type: 'string', description: '交易数据来源：CSV 文件路径，或直接给分录数组（每项 {date,desc,account,debit,credit}；列名中英皆可，自动跳过空科目行）' },
      output: { type: 'string', required: true, description: '报表落盘路径（.xlsx）。已存在时需显式 overwrite:true' },
      template: { type: 'string', description: '可选：以某份 xlsx 为底稿（保留其既有工作表与格式）。缺省新建空工作簿' },
      period: { type: 'string', description: '可选：只要该期间的记录，格式 YYYY-MM' },
      journal_sheet: { type: 'string', description: '日记账工作表名（缺省「日记账」）' },
      ledger_sheet: { type: 'string', description: '总账工作表名（缺省「总账」）' },
      pivot_sheet: { type: 'string', description: '透视表工作表名（缺省「透视表」）' },
      ...MODE_PARAMS,
    },
    execute: (args, ctl) => executeRunCode(CODE_ACCOUNTING_REPORT, argsData(args, ctl)),
  })

  reg(ctx, {
    name: 'office_check_workbook',
    description: '【任务级】动手前的体检（只读，不改任何东西）：列出每张工作表的已用区域/透视表/查询表/是否保护，标出**所有求值出错的公式单元格**、外部链接、定义名称、是否只读、是否有未保存改动，并给出 safe_to_edit 与逐条 flags。改工作簿之前先跑这个。',
    parameters: {
      path: { type: 'string', required: true, description: '工作簿完整路径' },
      sheet: { type: 'string', description: '只检查某张工作表（可选，缺省全部）' },
    },
    execute: (args, ctl) => executeRunCode(CODE_CHECK_WORKBOOK, argsData(args, ctl)),
  })

  reg(ctx, {
    name: 'office_replace_document_terms',
    description: '【任务级】Word 文档批量替换多组术语：一次给一组（[{find,replace}] 或 {find: replace}），先**全部计数**再统一替换，逐条报命中数与首处上下文，只动文本、保留原文格式。合同/模板里的甲方乙方、金额日期这类批量改写用这个，不要一串 word_edit 手动编排。mode=preview 只回报每组命中数，不改文档。',
    parameters: {
      path: { type: 'string', description: 'Word 文档完整路径（可选，缺省用活动文档）' },
      terms: {
        required: true,
        description: '替换术语：数组 [{find, replace}] 或对象 {find: replace}。replace 缺省为删除',
      },
      match_case: { type: 'boolean', required: false, description: '是否区分大小写（默认 true）。默认字面匹配，理由同 word_edit' },
      ...MODE_PARAMS,
    },
    execute: (args, ctl) => executeRunCode(CODE_REPLACE_TERMS, argsData(args, ctl)),
  })

  reg(ctx, {
    name: 'office_update_monthly_report',
    description: '【任务级】按期间更新月度报告：只替换该期间的行，**其他期间一行不动**（同一个期间重跑是替换不是追加，幂等）。原报告只读，结果写到 output。工作表须是月度表（含「期间」列）；表里是别的内容会报 OVERWRITE_NOT_CONFIRMED，不会默默清掉。mode=preview 回报该期间已有几行会被替换、其他期间保留几行、来料条数与借贷合计。',
    parameters: {
      report: { type: 'string', required: true, description: '要更新的月度报告路径（原文件不被修改）' },
      output: { type: 'string', required: true, description: '结果落盘路径。已存在时需显式 overwrite:true' },
      period: { type: 'string', required: true, description: '要更新的期间，格式 YYYY-MM' },
      source: { type: 'string', description: '该期间的数据来源：CSV 路径或分录数组（列名中英皆可）' },
      sheet: { type: 'string', description: '月度数据工作表名（缺省「月度数据」）' },
      total_sheet: { type: 'string', description: '可选：同时回报另一张总览表的行数（缺省「总览」）' },
      ...MODE_PARAMS,
    },
    execute: (args, ctl) => executeRunCode(CODE_UPDATE_MONTHLY, argsData(args, ctl)),
  })

  reg(ctx, {
    name: 'office_apply_template',
    description: '【任务级】把模板工作表的**外观**套到目标工作簿：字体/颜色/边框/数字格式（一次 PasteSpecial 搬完）、列宽、行高、冻结窗格。**目标表的数据一行不动**——数据与外观分家，才不会出现"套模板把数据弄丢"。结果写到 output，模板与目标原文件都不改。mode=preview 回报每对表的映射关系（模板表 → 目标表、目标表是否已存在）。',
    parameters: {
      template: { type: 'string', required: true, description: '模板工作簿路径' },
      target: { type: 'string', required: true, description: '要套格式的工作簿路径（原文件不被修改）' },
      output: { type: 'string', required: true, description: '结果落盘路径。已存在时需显式 overwrite:true' },
      sheet_map: { type: 'object', description: '表名映射 {模板表名: 目标表名}；缺省按同名对应，目标缺失则新建' },
      ...MODE_PARAMS,
    },
    execute: (args, ctl) => executeRunCode(CODE_APPLY_TEMPLATE, argsData(args, ctl)),
  })

  reg(ctx, {
    name: 'office_prepare_management_summary',
    description: '【任务级】从数据工作簿出一张「管理层摘要」：按维度聚合度量、排行取前 N、算占比，含「期间」列时还能给上期→本期对比。只用表里的真实数字算，不编造结论。group_by/measure 不给就自动认列（科目/部门/项目、借方/金额）。结果写到 output。mode=preview 回报分组数、合计与 TopN 预览。',
    parameters: {
      source: { type: 'string', required: true, description: '数据工作簿路径' },
      output: { type: 'string', required: true, description: '结果落盘路径。已存在时需显式 overwrite:true' },
      sheet: { type: 'string', description: '数据工作表名（可选，缺省活动表）' },
      group_by: { type: 'string', description: '分组列名（可选，缺省自动认「科目/部门/项目」）' },
      measure: { type: 'string', description: '度量列名（可选，缺省自动认「借方/金额」）' },
      top_n: { type: 'number', description: '摘要里列前几名（缺省 10）' },
      compare_period: { type: 'string', description: '可选：期间对比的当期标签（需表里有「期间」列）' },
      summary_sheet: { type: 'string', description: '摘要工作表名（缺省「管理层摘要」）' },
      ...MODE_PARAMS,
    },
    execute: (args, ctl) => executeRunCode(CODE_MGMT_SUMMARY, argsData(args, ctl)),
  })

  // ── kb_*：陌生目录的自建分类仓库（v1.2.0，工单 2026-10-03）─────────────────
  // 目标不是给 vault 加功能，而是让 agent 在陌生环境里靠 kb_scan_root 自己长出分类体系。
  // 身份哈希与 knowlp 侧 pool_scan.fingerprint_of 逐字同源（sha256(rel_posix \0 size \0 mtime_ns)）；
  // 词表与 pool_scan.POOL_EXTENSIONS / NOT_MATERIAL_EXTENSIONS 人工对齐——跨仓语言不同，
  // 两份清单改动时必须两边一起改（docs/kb.md 对齐纪律）。
  // 红线：只读（v1 无任何 move/rename/delete 代码路径）；状态产物只落状态目录，绝不写被测目录。

  const KB_POOL_EXTENSIONS = {
    text: ['.md', '.txt', '.markdown'],
    pdf: ['.pdf'],
    image: ['.png', '.jpg', '.jpeg', '.gif', '.webp', '.bmp', '.svg', '.heic', '.ico', '.tiff'],
    office: ['.xlsx', '.xls', '.docx', '.doc', '.pptx', '.ppt', '.csv', '.one', '.et', '.wps'],
    code: ['.py', '.js', '.ts', '.mjs', '.cjs', '.json', '.yaml', '.yml', '.toml', '.sh',
           '.bat', '.ps1', '.sql', '.html', '.css', '.ipynb', '.r', '.java', '.go', '.rs',
           '.cmd', '.ini', '.pem', '.tag', '.nu', '.sample', '.example'],
    video: ['.mp4', '.mov', '.avi', '.mkv', '.webm', '.flv', '.wmv'],
  }
  // 「本就不是资料」档（P1.5-2）：派生物/可执行/快捷方式/VCS 内部件——登记但不进任何下游
  const KB_NOT_MATERIAL = ['.pyc', '.exe', '.dll', '.lnk', '.url', '.rev', '.pack', '.idx', '.msi', '.class']
  const KB_MAGIC = [ // [首字节 hex, pool]——blob 侧 bytes.fromhex 还原（JSON 不携带裸字节）
    ['25504446', 'pdf'], ['89504e47', 'image'], ['ffd8ff', 'image'], ['47494638', 'image'],
    ['504b0304', 'mixed'], ['1f8b', 'unknown'], ['494433', 'video'], ['66747970', 'video'],
    ['52617221', 'unknown'], ['377abcaf271c', 'unknown'],
  ]

  /** kb blob：自包含 Python——不经 comCode 的工作簿前置（kb 一级扫描不开工作簿也能跑）。
   *  枚举走 FSO（工单钦定通道）；身份字段 size/mtime_ns 取 os.stat 权威值——FSO 的
   *  DateLastModified 只有秒精度，会让指纹与 pool_scan 失配，所以 FSO 只管枚举。 */
  const kbCode = (body) => [
    "import json, hashlib, os, stat, sys",
    "args = json.loads(data) if data else {}",
    "try:",
    "    import pythoncom",
    "    from win32com.client import Dispatch, DispatchEx",
    "    pythoncom.CoInitialize()",
    "except Exception as _e:",
    "    raise Exception('[APP_UNAVAILABLE] COM 通道不可用: %s' % _e)",
    "try:",
    ...body.map((l) => (l ? '    ' + l : l)),
    "except Exception as _e:",
    "    _m = str(_e)",
    "    if _m.startswith('['):",
    "        raise                                # 🟠-7: 内层已带稳定码（OPEN_FAILED 等），不再叠加前缀",
    "    raise Exception('[KB_SCAN_FAILED] %s' % _m)",
  ].join('\n')

  const KB_CODE = kbCode([
    "root = os.path.realpath(args['root'])",
    'max_files = int(args.get("max_files") or 20000)',
    'max_deep = int(args.get("max_deep") or 0)',
    'vocab = args.get("vocab") or {}',
    'not_material = set(args.get("not_material") or [])',
    'magic = [(bytes.fromhex(h), p) for h, p in (args.get("magic") or [])]',
    'if not os.path.isdir(root):',
    "    raise Exception('[OPEN_FAILED] 目录不存在: %s' % root)",
    "fso = Dispatch('Scripting.FileSystemObject')",
    'def _sniff(path):',
    '    # magic(首16字节) + 文本性(4KB 无 NUL)——与 knowlp 侧 pool_scan.classify 同判据',
    '    try:',
    '        with open(path, "rb") as f: chunk = f.read(4096)',
    '    except Exception:',
    '        return None, True',
    '    for m, p in magic:',
    '        if chunk[:len(m)] == m: return p, False',
    '    if chunk and b"\\x00" not in chunk: return "text", False',
    '    return None, False',
    'entries, unrec, notmat = {}, [], []',
    'seen_dirs = set([os.path.realpath(root)])',
    'count, dup, truncated, skipped_dot = 0, 0, False, 0',
    'stack = [(root, "")]',
    'while stack and not truncated:',
    '    d, rel_d = stack.pop()',
    '    try:',
    '        folder = fso.GetFolder(d)',
    '        subs = [s.Name for s in folder.SubFolders]',
    '        files = [fi.Name for fi in folder.Files]',
    '    except Exception:',
    '        continue',
    '    for n in sorted(subs):',
    '        if n.startswith("."):',
    '            skipped_dot += 1',
    '            continue                       # 点目录整棵跳过（.git/.obsidian/...）——既有规则，docs/kb.md 写明',
    '        full = os.path.join(d, n)',
    '        rp = os.path.realpath(full)',
    '        if rp != root and not rp.startswith(root + os.sep): continue   # ★12 越界保险丝：realpath 不在 root 下 → 整棵剪掉（与链接类型无关）',
    '        try:',
    '            lst = os.lstat(full)',
    '        except Exception:',
    '            continue',
    '        if stat.S_ISLNK(lst.st_mode) or getattr(lst, "st_reparse_tag", 0) == 0xA0000003: continue   # ★12 junction/symlink 目默认不追——junction 是 reparse 点不是 symlink，is_symlink 挡不住（2026-10-06 真实复现）',
    '        if rp in seen_dirs: continue       # 链接环守卫',
    '        seen_dirs.add(rp)',
    '        stack.append((full, ((rel_d + "/") if rel_d else "") + n))',
    '    for n in sorted(files):',
    '        if n.startswith("."):',
    '            skipped_dot += 1               # 点文件同样整类跳过（.env/.gitignore/...）',
    '            continue',
    '        full = os.path.join(d, n)',
    '        try:',
    '            lst = os.lstat(full)',
    '        except Exception:',
    '            lst = None',
    '        if lst is not None and (stat.S_ISLNK(lst.st_mode) or getattr(lst, "st_reparse_tag", 0) == 0xA0000003):',
    '            continue                       # ★12 文件级链接默认不登记、不占配额（OneDrive 等云占位 tag 不在此列，照常登记）',
    '        if max_files and count >= max_files:',
    '            truncated = True',
    '            break',
    '        count += 1                             # 🟠-5: 只数真正登记的文件，点文件不吃配额',
    '        rel_posix = ((rel_d + "/") if rel_d else "") + n',
    '        try:',
    '            st = os.stat(full)',
    '        except Exception:',
    '            unrec.append({"source_uri": "file://" + rel_posix, "reason": "unreadable", "ext": os.path.splitext(n)[1].lower().lstrip(".")})',
    '            continue',
    '        fp = hashlib.sha256(("%s\\x00%d\\x00%d" % (rel_posix, st.st_size, st.st_mtime_ns)).encode("utf-8")).hexdigest()',
    '        ext = os.path.splitext(n)[1].lower()',
    '        fmt = ext.lstrip(".")',
    '        base = {"source_uri": "file://" + rel_posix, "format": fmt or None,',
    '                "size_bytes": st.st_size, "mtime_epoch": st.st_mtime, "fingerprint": fp}',
    '        if fp in entries:',
    '            dup += 1',
    '            continue',
    '        if ext and ext in not_material:',
    '            e = dict(base, pool="not-material", tier="not-material", classified_by="extension", reason="not-material-by-extension")',
    '            entries[fp] = e',
    '            notmat.append(e)',
    '            continue',
    '        pool_claim = None',
    '        for p_name, exts in vocab.items():',
    '            if ext and ext in exts: pool_claim = p_name',
    '        if pool_claim:',
    '            entries[fp] = dict(base, pool=pool_claim, classified_by="extension")',
    '            continue',
    '        sniff_pool, unreadable = _sniff(full)   # 🟠-4: 未认领/无扩展名都回落 magic（改名 .dat 的 PDF 也判得出）',
    '        if sniff_pool is None:',
    '            reason = "unreadable" if unreadable else ("no-extension" if not ext else "unclaimed-extension")',
    '            e = dict(base, pool="unknown", classified_by="unknown", reason=reason)',
    '            entries[fp] = e',
    '            unrec.append(e)',
    '            continue',
    '        if sniff_pool == "unknown":            # 🟠-3: magic 认出容器但无池可归（gzip/rar/7z）→ 显式进桶',
    '            e = dict(base, pool="unknown", classified_by="magic", reason="unknown-magic")',
    '            entries[fp] = e',
    '            unrec.append(e)',
    '            continue',
    '        entries[fp] = dict(base, pool=sniff_pool, classified_by="magic")',
    'deep = []',
    'if max_deep > 0:',
    '    sheet_exts = set(args.get("deep_exts") or [".xlsx", ".xlsm", ".xls"])',
    '    cands = [e for e in entries.values()',
    '             if e.get("pool") == "office" and ("." + (e.get("format") or "")) in sheet_exts][:max_deep]',
    '    # 容器预检：非法容器绝不递给 Excel——截断/伪造的 xlsx 会触发修复对话框，挂死整条通道。',
    '    # .xlsx/.xlsm 必须 PK\\x03\\x04 开头，.xls 必须 OLE2 头（\\xd0\\xcf\\x11\\xe0）。',
    '    def _valid_container(path):',
    '        try:',
    '            with open(path, "rb") as f: head = f.read(8)',
    '        except Exception:',
    '            return False',
    '        if path.lower().endswith(".xls") and not path.lower().endswith((".xlsx", ".xlsm")):',
    '            return head[:4] == b"\\xd0\\xcf\\x11\\xe0"',
    '        return head[:4] == b"PK\\x03\\x04"',
    '    valid_cands = []',
    '    for e in cands:',
    '        path = os.path.join(root, e["source_uri"].split("://", 1)[1])   # 🔴-1: source_uri 是 root 相对路径，必须拼回 root',
    '        if _valid_container(path):',
    '            valid_cands.append(e)',
    '        else:',
    '            deep.append({"source_uri": e["source_uri"], "error": "not-a-valid-workbook-container"})',
    '    cands = valid_cands',
    '    excel = None',
    '    if cands:',
    '        try:',
    '            excel = DispatchEx("Excel.Application")',
    '            excel.Visible = False',
    '            excel.DisplayAlerts = False',
    '        except Exception as _de:',
    '            deep.append({"error": "excel-unavailable: %s" % _de})   # 🔴-2: 无 Excel 的机器上一级结果不能陪葬',
    '    opened = []',
    '    try:',
    '        for e in cands:',
    '            if excel is None: break   # 🔴-2: 起不来就不扫，一级结果照常返回',
    '            path = os.path.join(root, e["source_uri"].split("://", 1)[1])   # 🔴-1: source_uri 是 root 相对路径，必须拼回 root',
    '            rec = {"source_uri": e["source_uri"]}',
    '            try:',
    '                wb = excel.Workbooks.Open(path, 0, True)   # UpdateLinks=0, ReadOnly=True',
    '                opened.append(wb)',
    '                sheets = []',
    '                for ws in wb.Worksheets:',
    '                    try: used = str(ws.UsedRange.Address)',
    '                    except Exception: used = None',
    '                    header = []',
    '                    try:',
    '                        vals = ws.Rows(1).Value',
    '                        flat = vals[0] if isinstance(vals, tuple) and vals and isinstance(vals[0], tuple) else vals',
    '                        if isinstance(flat, tuple): header = [str(v) for v in flat][:12]',
    '                        elif flat is not None: header = [str(flat)]',
    '                    except Exception: pass',
    '                    try: hl = int(ws.Hyperlinks.Count)',
    '                    except Exception: hl = 0',
    '                    sheets.append({"name": ws.Name, "used_range": used, "header": header, "hyperlinks": hl})',
    '                try: links = [str(x) for x in (wb.LinkSources(1) or [])]',
    '                except Exception: links = []',
    '                rec["sheets"] = sheets',
    '                rec["link_sources"] = links',
    '            except Exception as we:',
    '                rec["error"] = str(we)',
    '            deep.append(rec)',
    '    finally:',
    '        for wb in opened:',
    '            try: wb.Close(False)',
    '            except Exception: pass',
    '        if excel is not None:',
    '            try: excel.Quit()',
    '            except Exception: pass',
    'vba_channel = "VBA_ACCESS_DENIED"',
    'try:',
    '    import winreg',
    '    for ver in ("16.0", "15.0", "14.0"):',
    '        try:',
    '            k = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Software\\\\Microsoft\\\\Office\\\\%s\\\\Excel\\\\Security" % ver)',
    '        except OSError:',
    '            continue',
    '        try:',
    '            val, _ = winreg.QueryValueEx(k, "AccessVBOM")',
    '            if int(val) == 1: vba_channel = "available"',
    '        except OSError: continue',
    '        break                              # 🟠-6: 读到值才停（键存在但未设值 → 继续试旧版本）',
    'except Exception: pass',
    "output = json.dumps({'root': root, 'stats': {'files_registered': len(entries), 'duplicate_identities': dup, 'skipped_dot': skipped_dot, 'truncated': truncated}, 'entries': entries, 'unclassified': {'unrecognized': unrec, 'not_material': notmat}, 'deep': deep, 'vba_channel': vba_channel}, ensure_ascii=False, default=str)",
  ])

  /** kb 写回 blob（第九批 ★6）：**只写自定义文档属性，不碰内容**——命门。
   *  三个动作：write（双闸后才到这）/ rollback（删本次写入的 kb_* 属性）/ read（W2 读回核对）。
   *  一致性判据是**语义等价**不是字节：写回前后各拍一次内容快照（Excel=逐格 值+公式，
   *  Word=正文文本，PPT=文本框文本），sha256 比对；属性闭环=关掉重开（只读）再读一遍。
   *  preview 用只读打开——mtime 物理上不可能变。 */
  const kbwCode = (body) => [
    "import json, hashlib, os, sys, datetime",
    "args = json.loads(data) if data else {}",
    "if args.get('path'): args['path'] = args['path'].replace('/', '\\\\')",
    "try:",
    "    import pythoncom",
    "    from win32com.client import DispatchEx",
    "    pythoncom.CoInitialize()",
    "except Exception as _e:",
    "    raise Exception('[APP_UNAVAILABLE] COM 通道不可用: %s' % _e)",
    "try:",
    ...body.map((l) => (l ? '    ' + l : l)),
    "except Exception as _e:",
    "    _m = str(_e)",
    "    if _m.startswith('['):",
    "        raise                                # 内层已带稳定码（FILE_LOCKED 等），不叠加前缀",
    "    raise Exception('[KB_WRITE_FAILED] %s' % _m)",
  ].join('\n')

  const KBW_CODE = kbwCode([
    "path = args['path']",
    "if not os.path.isfile(path):",
    "    raise Exception('[OPEN_FAILED] 文件不存在: %s' % path)",
    "props = args.get('properties') or {}",
    "rollback = bool(args.get('rollback'))",
    "read_only_action = bool(args.get('read_only_action'))",
    "ctl = args.get('_ctl') or {}",
    "preview = bool(ctl.get('preview'))",
    "import re as _re",
    "for _k in props:",
    "    if not _re.match(r'^kb_[A-Za-z0-9_-]{1,60}$', _k):",
    "        raise Exception('[MISSING_PARAM] 属性名必须以 kb_ 开头且只含字母数字_-: %r' % _k)",
    "MAX_PROP_LEN = 1024",
    "def _j(x):",
    "    if isinstance(x, (tuple, list)): return [_j(i) for i in x]",
    "    if isinstance(x, (datetime.datetime, datetime.date)): return x.isoformat()",
    "    return x",
    "def _sha(snap):",
    "    return hashlib.sha256(json.dumps(snap, ensure_ascii=False, default=str, sort_keys=True).encode('utf-8')).hexdigest()",
    "def _snap(doc, kind):",
    "    if kind == 'excel':",
    "        out = []",
    "        for ws in doc.Worksheets:",
    "            ur = ws.UsedRange",
    "            try: vals = _j(ur.Value)",
    "            except Exception: vals = None",
    "            try: forms = _j(ur.Formula)",
    "            except Exception: forms = None",
    "            out.append([ws.Name, str(ur.Address), vals, forms])",
    "        return out",
    "    if kind == 'word': return _j(doc.Content.Text)",
    "    out = []",
    "    for sl in doc.Slides:",
    "        for sh in sl.Shapes:",
    "            try: has = bool(sh.HasTextFrame)",
    "            except Exception: has = False",
    "            out.append(_j(sh.TextFrame.TextRange.Text) if has else '')",
    "    return out",
    "def _read_props(doc):",
    "    # B3（DSH 验收 2026-10-06）：原实现 `cps.Item(i)` 在文件**已含自定义属性**时抛",
    "    # -2147352573 '找不到成员。'（DISP_E_MEMBERNOTFOUND）—— 动态派发下 DocumentProperties",
    "    # 的 Item 不经参数化属性调用；文件没有属性时 Count=0、循环不执行，所以这个 bug 只在",
    "    # '已经有属性' 时暴露（正是写回之后）。改按可靠性依次退：默认索引 cps(i) → _NewEnum 迭代。",
    "    # **两条都失败则抛错**，不返回空 dict —— 返回空会把'读不到'伪装成'没有属性'。",
    "    out = {}",
    "    cps = doc.CustomDocumentProperties",
    "    n = int(cps.Count)",
    "    if n == 0:",
    "        return out",
    "    items = None",
    "    try:",
    "        items = [cps(i) for i in range(1, n + 1)]",
    "    except Exception:",
    "        try:",
    "            items = list(cps)",
    "        except Exception as _e2:",
    "            raise Exception('[KB_PROP_READ_FAILED] 自定义属性读不到（索引与枚举都失败）: %s' % _e2)",
    "    for p in items:",
    "        try: out[str(p.Name)] = '' if p.Value is None else str(p.Value)",
    "        except Exception: pass",
    "    return out",
    "def _find_prop(d, name):",
    "    # B3 残留（DSH 复跑 2026-10-06）：`cps.Item(name)` 与 `cps.Item(i)` 同病 —— 动态派发下",
    "    # 具名 Item 在集合非空时抛 DISP_E_MEMBERNOTFOUND。删除与更新都必须先**按索引/枚举**",
    "    # 定位到对象，再对它操作。定位不到返回 None，由调用方决定是 Add 还是放弃。",
    "    try:",
    "        cps = d.CustomDocumentProperties",
    "        n = int(cps.Count)",
    "    except Exception:",
    "        return None",
    "    if n == 0: return None",
    "    items = None",
    "    try:",
    "        items = [cps(i) for i in range(1, n + 1)]",
    "    except Exception:",
    "        try: items = list(cps)",
    "        except Exception: return None",
    "    for p in items:",
    "        try:",
    "            if str(p.Name) == name: return p",
    "        except Exception: pass",
    "    return None",
    "def _lock_check(e):",
    "    m = str(e).lower()",
    "    if any(k in m for k in ('locked', '被占用', 'in use', 'sharing', '正在使用', '只读')):",
    "        raise Exception('[FILE_LOCKED] 文件被占用，稍后重试: %s' % str(e)[:160])",
    "ext = os.path.splitext(path)[1].lower()",
    "kind = 'excel' if ext in ('.xlsx', '.xlsm', '.xls') else ('word' if ext in ('.docx', '.doc') else ('ppt' if ext in ('.pptx', '.ppt') else None))",
    "if kind is None:",
    "    raise Exception('[MISSING_PARAM] 不支持的格式 %s（仅 xlsx/xlsm/xls、docx/doc、pptx/ppt）' % ext)",
    "app = None",
    "doc = None",
    "opened_path = None",
    "def _open(read_only):",
    "    # DispatchEx=独立实例：绝不附着用户正在用的 Office；整会话 Quit 回收（零残留）。",
    "    a = None",
    "    d = None",
    "    try:",
    "        if kind == 'excel':",
    "            a = DispatchEx('Excel.Application')",
    "            a.Visible = False; a.DisplayAlerts = False",
    "            d = a.Workbooks.Open(path, 0, read_only)",
    "        elif kind == 'word':",
    "            a = DispatchEx('Word.Application')",
    "            a.Visible = False; a.DisplayAlerts = False",
    "            d = a.Documents.Open(path, False, read_only)",
    "        else:",
    "            a = DispatchEx('PowerPoint.Application')",
    "            d = a.Presentations.Open(path, read_only, False, False)   # ReadOnly, Untitled, WithWindow=False（不弹窗口）",
    "    except Exception as e:",
    "        try:",
    "            if a is not None: a.Quit()",
    "        except Exception: pass",
    "        _lock_check(e)",
    "        raise Exception('[KB_WRITE_FAILED] 打开失败: %s' % str(e)[:200])",
    "    return a, d",
    "def _shut(d, a, save):",
    "    try:",
    "        if d is not None:",
    "            if kind == 'excel': d.Close(save)",
    "            elif kind == 'word': d.Close(save)",
    "            else: d.Close()",
    "    except Exception: pass",
    "    try:",
    "        if a is not None: a.Quit()",
    "    except Exception: pass",
    "def _persist(d):",
    "    try:",
    "        if kind == 'ppt': d.Save()",
    "        else: d.Save()",
    "    except Exception as e:",
    "        _lock_check(e)",
    "        raise Exception('[SAVE_FAILED] 保存失败: %s' % str(e)[:200])",
    "if preview or read_only_action:",
    "    a, d = _open(True)                       # 只读打开：preview/read 物理上不可能落笔",
    "    try:",
    "        snap_sha = _sha(_snap(d, kind))",
    "        props_now = _read_props(d)",
    "    finally:",
    "        _shut(d, a, False)",
    "    if read_only_action:",
    "        output = json.dumps({'action': 'read', 'properties': props_now, 'content_sha256': snap_sha, 'content_kind': kind}, ensure_ascii=False, default=str)",
    "    else:",
    "        plan = {'preview': True, 'tool': 'kb_write_back_office', 'action': 'rollback' if rollback else 'write',",
    "                'path': path, 'content_kind': kind, 'content_sha256_before': snap_sha,",
    "                'properties': [{'name': k, 'exists': k in props_now, 'current': props_now.get(k),",
    "                                'next': ('(delete)' if rollback else str(props[k])[:MAX_PROP_LEN])} for k in props],",
    "                'changed': False, 'saved': False}",
    "        output = json.dumps(plan, ensure_ascii=False, default=str)",
    "else:",
    "    a, d = _open(False)",
    "    try:",
    "        snap_before = _sha(_snap(d, kind))",
    "        props_before = _read_props(d)",
    "        if rollback:",
    "            applied = []",
    "            for k in props:",
    "                if k in props_before:",
    "                    _p = _find_prop(d, k)",
    "                    if _p is not None:",
    "                        try: _p.Delete()",
    "                        except Exception as _de: raise Exception('[KB_WRITE_FAILED] 删属性失败 %s: %s' % (k, _de))",
    "                    applied.append({'name': k, 'op': 'deleted'})",
    "            _persist(d)",
    "        else:",
    "            if not ctl.get('confirm'):",
    "                raise Exception('[RISKY_OP_NOT_CONFIRMED] 写回需要 confirm:true（双闸，缺一不写）')",
    "            applied = []",
    "            for k, v in props.items():",
    "                sval = str(v)",
    "                truncated = len(sval) > MAX_PROP_LEN",
    "                if truncated: sval = sval[:MAX_PROP_LEN]",
    "                _p = _find_prop(d, k)",
    "                if _p is not None:",
    "                    # 已存在 → 改它。原实现 `Item(k).Value=` 抛错后兜到 .Add()，而对**已存在**",
    "                    # 的属性 .Add() 必然失败（实测 `-2147352567 发生意外`）。改失败就响亮报错。",
    "                    try: _p.Value = sval",
    "                    except Exception as _ue: raise Exception('[KB_WRITE_FAILED] 改属性失败 %s: %s' % (k, _ue))",
    "                    applied.append({'name': k, 'op': 'updated', 'truncated': truncated})",
    "                else:",
    "                    d.CustomDocumentProperties.Add(k, False, 4, sval)",
    "                    applied.append({'name': k, 'op': 'added', 'truncated': truncated})",
    "            _persist(d)",
    "    finally:",
    "        _shut(d, a, False)",
    "    # 关掉重开（只读）核账：属性真的持久化了 + 内容快照逐位一致（语义等价判据）",
    "    a2, d2 = _open(True)",
    "    try:",
    "        props_after = _read_props(d2)",
    "        snap_after = _sha(_snap(d2, kind))",
    "    finally:",
    "        _shut(d2, a2, False)",
    "    if rollback:",
    "        leftovers = [k for k in props if k in props_after]",
    "        if leftovers: raise Exception('[KB_WRITE_FAILED] 回滚后属性仍存在: %s' % ','.join(leftovers))",
    "    else:",
    "        mismatched = [k for k, v in props.items() if props_after.get(k) != str(v)[:MAX_PROP_LEN]]",
    "        if mismatched: raise Exception('[KB_WRITE_FAILED] 属性读回不一致: %s' % ','.join(mismatched))",
    "    if snap_after != snap_before:",
    "        raise Exception('[KB_WRITE_FAILED] 内容快照不一致——写回碰了不该碰的东西，立即检查')",
    "    output = json.dumps({'action': 'rollback' if rollback else 'write', 'path': path,",
    "                         'content_kind': kind, 'applied': applied,",
    "                         'properties_before': props_before, 'properties_after': props_after,",
    "                         'content_sha256_before': snap_before, 'content_sha256_after': snap_after,",
    "                         'data_identical': True}, ensure_ascii=False, default=str)",
  ])


  /** Obsidian 硬守卫（红线 2）：root 是 vault、含 vault、或位于 vault 内 → 一律拒绝。
   *  纯 JS，先于 COM 通道执行——降级态下守卫依然生效。 */
  const kbVaultGuard = (rootReal) => {
    const v = process.env.KNOWLP_VAULT
    if (!v) return { vault: null, guarded: false }
    let vaultReal
    try {
      vaultReal = realpathSync.native(v)
    } catch {
      return { vault: v, guarded: false }
    }
    const norm = (p) => String(p).toLowerCase().replace(/[\\/]+$/, '')
    const a = norm(rootReal)
    const b = norm(vaultReal)
    const overlap = a === b || a.startsWith(b + '\\') || a.startsWith(b + '/') || b.startsWith(a + '\\') || b.startsWith(a + '/')
    return { vault: vaultReal, guarded: overlap }
  }

  // ── kb_taxonomy_*（P2-a）：agent 自己长出来的分类学 ─────────────────
  // 纯 JS 状态工具（无 COM）：taxonomy.json 落 kb/<root_id>/，与 registry 同域。
  // 分类不是固定枚举——增/改名/合并/删由 agent 驱动；supersedes 沿用 evidence.py 语义。
  // 漂移审计：重叠分类（同 kind+pattern 出现在两个类别）→ 审计退出码 1。

  const KB_RULE_KINDS = ['ext', 'name-contains', 'path-contains']

  const kbTaxonomyPath = (rootReal) => {
    const rootId = createHash('sha256').update(rootReal).digest('hex').slice(0, 16)
    return join(STATE_DIR, 'kb', rootId, 'taxonomy.json')
  }

  const kbLoadTaxonomy = (path) => {
    try {
      const t = JSON.parse(readFileSync(path, 'utf-8'))
      if (!Array.isArray(t.categories)) throw new Error('categories missing')
      return t
    } catch {
      return { version: 0, categories: [] }
    }
  }

  /** 漂移审计：返回问题清单；空数组 = 干净。
   *  1) 重叠分类：同 (kind, pattern) 规则出现在两个类别
   *  2) 环：parent 链成环
   *  3) 悬空：parent_id / supersedes 指向不存在的类别 */
  const kbAuditTaxonomy = (tax) => {
    const problems = []
    const byId = new Map(tax.categories.map((c) => [c.id, c]))
    const seen = new Map()
    for (const c of tax.categories) {
      for (const r of c.rules || []) {
        const key = (r.kind || '') + '::' + (r.pattern || '')
        if (seen.has(key)) {
          problems.push({ kind: 'overlap', rule: r, categories: [seen.get(key), c.id] })
        } else {
          seen.set(key, c.id)
        }
        if (r.kind && !KB_RULE_KINDS.includes(r.kind)) {
          problems.push({ kind: 'bad-rule-kind', rule: r, category: c.id })
        }
      }
      if (c.parent_id && !byId.has(c.parent_id)) {
        problems.push({ kind: 'dangling-parent', category: c.id, parent_id: c.parent_id })
      }
      for (const sup of c.supersedes || []) {
        if (!byId.has(sup)) {
          problems.push({ kind: 'dangling-supersedes', category: c.id, supersedes: sup })
        }
      }
    }
    for (const c of tax.categories) {
      const chain = new Set([c.id])
      let cur = c
      while (cur?.parent_id) {
        if (chain.has(cur.parent_id)) {
          problems.push({ kind: 'parent-cycle', category: c.id })
          break
        }
        cur = byId.get(cur.parent_id)
      }
    }
    return problems
  }

  const kbTaxonomyGet = (args) => {
    let rootReal
    try {
      rootReal = realpathSync.native(args.root)
    } catch {
      return { ok: false, error: '[OPEN_FAILED] 目录不存在: ' + args.root }
    }
    const path = kbTaxonomyPath(rootReal)
    const tax = kbLoadTaxonomy(path)
    const problems = kbAuditTaxonomy(tax)
    return {
      ok: true,
      output: { root: rootReal, taxonomy_path: path, taxonomy: tax,
                audit: { problems, clean: problems.length === 0 } },
    }
  }

  const kbTaxonomySet = (args) => {
    let rootReal
    try {
      rootReal = realpathSync.native(args.root)
    } catch {
      return { ok: false, error: '[OPEN_FAILED] 目录不存在: ' + args.root }
    }
    const ops = args.ops
    if (!Array.isArray(ops) || !ops.length) {
      return { ok: false, error: '[MISSING_PARAM] ops 必须是非空数组（add/rename/merge/delete）' }
    }
    const path = kbTaxonomyPath(rootReal)
    const tax = kbLoadTaxonomy(path)
    const applied = []
    for (const op of ops) {
      const kind = op.op
      if (kind === 'add') {
        const c = op.category || {}
        if (!c.id || !c.label) {
          return { ok: false, error: '[MISSING_PARAM] add 需要 category.id 与 category.label' }
        }
        if (tax.categories.some((x) => x.id === c.id)) {
          return { ok: false, error: `[DUPLICATE_ID] 类别已存在: ${c.id}` }
        }
        tax.categories.push({
          id: c.id, label: c.label, parent_id: c.parent_id || null,
          description: c.description || '',
          rules: (c.rules || []).map((r) => ({ kind: r.kind, pattern: r.pattern })),
          created_at: new Date().toISOString(),
        })
        applied.push({ op: 'add', id: c.id })
      } else if (kind === 'rename') {
        const c = tax.categories.find((x) => x.id === op.id)
        if (!c) return { ok: false, error: `[CATEGORY_NOT_FOUND] ${op.id}` }
        c.label = op.label || c.label
        c.description = op.description ?? c.description
        applied.push({ op: 'rename', id: op.id })
      } else if (kind === 'merge') {
        const into = tax.categories.find((x) => x.id === op.into_id)
        if (!into) return { ok: false, error: `[CATEGORY_NOT_FOUND] ${op.into_id}` }
        for (const fid of op.from_ids || []) {
          const from = tax.categories.find((x) => x.id === fid)
          if (!from) continue
          into.rules = [...(into.rules || []), ...(from.rules || [])]
          into.supersedes = [...new Set([...(into.supersedes || []), from.id,
            ...(from.supersedes || [])])]
          tax.categories = tax.categories.filter((x) => x.id !== fid)
          applied.push({ op: 'merge', from: fid, into: op.into_id })
        }
      } else if (kind === 'delete') {
        const hasChildren = tax.categories.some((x) => x.parent_id === op.id)
        if (hasChildren && !op.force) {
          return { ok: false, error: `[HAS_CHILDREN] ${op.id} 有子类别——先合并/删除子类别，或 force:true` }
        }
        tax.categories = tax.categories.filter((x) => x.id !== op.id)
        applied.push({ op: 'delete', id: op.id })
      } else {
        return { ok: false, error: `[UNKNOWN_OP] ${kind}（应为 add/rename/merge/delete）` }
      }
    }
    tax.version = (tax.version || 0) + 1
    const problems = kbAuditTaxonomy(tax)
    mkdirSync(join(STATE_DIR, 'kb', createHash('sha256').update(rootReal).digest('hex').slice(0, 16)), { recursive: true })
    writeFileSync(path, JSON.stringify(tax, null, 1), 'utf-8')
    return {
      ok: true,
      output: { root: rootReal, taxonomy_path: path, applied, version: tax.version,
                categories: tax.categories.length,
                audit: { problems, clean: problems.length === 0 } },
    }
  }

  // ── B2（DSH 验收 2026-10-06）：嵌套派发的输入形状 ─────────────────────────
  // dsh-tools :1298-1308 的形状；两个 seam（语义缝 / 写回缝）共用这一处构造，避免再漏字段。
  //   · signal   —— **必须带**：dsh-tools:3161 读 `signal.aborted`，缺了抛 TypeError（本次实测）
  //   · parent   —— 标它是嵌套派发（:3141）；取自外层 exec.token
  //   · callId   —— 必须唯一，否则 host 当成同一个调用
  // 没有外层 exec 时（单测直调 / 非 dsh 宿主）退化为一个**永不中止的新信号** ——
  // 满足形状、不至于崩，但不具备调用方取消传播（拿不到就如实退化，不假装有）。
  let kbNestedSeq = 0
  const kbNestedCall = (exec, name, args) => ({
    callId: `${exec?.callId != null ? String(exec.callId) : 'kb'}:n${++kbNestedSeq}`,
    rootCallId: exec?.rootCallId ?? exec?.callId,
    name,
    arguments: args,
    ...(exec?.agent ? { agent: exec.agent } : {}),
    ...(exec?.token ? { parent: exec.token } : {}),
    signal: exec?.signal ?? new AbortController().signal,
  })

  reg(ctx, {
    name: 'kb_classify',
    description: '【kb · 只读】用分类学规则（ext / name-contains / path-contains）把 registry 条目归到类别：单命中归档、零命中报未归类、多命中报冲突。冻结桶（not-material / 带 reason）不参与。use_semantics=true 走 KnowLP 语义缝三级梯子（approved 审批派发 → direct 降级 → unavailable 如实声明）——seam 断链或失败时显式降级并带错误，绝不假装做了语义判断。',
    parameters: {
      root: { type: 'string', required: true, description: '已扫描并已建分类学的目录' },
      // §1.8-1 修正（DSH 验收 2026-10-06）：这个开关原先没进 schema，agent 只读 schema
      // 根本发现不了它 ⇒ 语义缝永远没人调用。补上。
      use_semantics: { type: 'boolean', description: 'true = 对待归类项走 KnowLP 语义缝三级梯子（approved 审批派发 → direct 降级 → unavailable 如实声明）。缺省 false = 只做机械规则。返回的 semantics 说明走到了哪一级，seam_how 说明工具名是怎么解析到的' },
    },
    execute: async (args, ctl, exec) => kbClassifyTool(args, exec),
  })

  reg(ctx, {
    name: 'kb_scheme_propose',
    description: '【kb · 只读】产影子目录方案：按分类学的 label 链生成每个文件的影子路径（命名策略：原文件名，重名加指纹前 8 位）。方案落状态目录 scheme-<plan_id>.json，**不碰任何文件**；unassigned/conflicts/冻结条目单独计数，不进方案。',
    parameters: {
      root: { type: 'string', required: true, description: '已扫描并已建分类学的目录' },
    },
    execute: async (args) => kbSchemePropose(args),
  })

  reg(ctx, {
    name: 'kb_annotate',
    description: '【kb · 变更】给已登记的资料写标注：annotations.jsonl 每行是 ContextItem 形状 + 标注字段（登记优先）。文件已不在或指纹已变 → unverifiable 如实标记；写回原文件（.md 前置元数据 / Office 自定义文档属性）默认关——write_back:true && confirm:true 双闸。Office 写回**只写 kb_* 属性、绝不碰单元格/正文/幻灯片**，且**只走审批通道**（拿不到就不写，只登记）。',
    parameters: {
      root: { type: 'string', required: true, description: '资料所在目录（必须已 kb_scan_root）' },
      fingerprint: { type: 'string', required: true, description: 'registry 里的文件指纹' },
      note: { type: 'string', description: '标注正文' },
      tags: { type: 'array', description: '标签（字符串数组）' },
      taxonomy_id: { type: 'string', description: '归属类别 id（可选）' },
      sensitivity: { type: 'string', description: 'public | private | commercial（缺省 private——缺省禁止云回退，放行需显式打 public）' },
      author: { type: 'string', description: '标注者（agent 名或用户名）' },
      require_intact: { type: 'boolean', description: 'true = 文件已不在/指纹已变时直接拒绝（缺省只登记并标 unverifiable）' },
      write_back: { type: 'boolean', description: '写回原文件（默认关；.md 写前置元数据，Office 写 kb_* 自定义文档属性）' },
      confirm: { type: 'boolean', description: '写回确认位（与 write_back 同时 true 才生效）' },
    },
    execute: async (args, ctl, exec) => kbAnnotateImpl(args, {
      // W4（工单 2026-10-06 第九批）：Office 写回经**审批通道**派发 kb_write_back_office。
      // 与语义缝的梯子不同 —— **写入类不做 direct 降级**：直接 execute 一个工具对象会绕过
      // 审批，对写动作不可接受。拿不到审批通道就返回 error，由调用方降级为「只登记」。
      officeWriter: async (absPath, properties) => {
        const t = _ctx?.tools
        if (!(t && typeof t.execute === 'function' && typeof t.get === 'function' &&
              t.get('kb_write_back_office'))) {
          return { error: 'seam-unavailable：拿不到 kb_write_back_office 的审批通道' }
        }
        const r = await t.execute(kbNestedCall(exec, 'kb_write_back_office',
          { path: absPath, properties, mode: 'managed', confirm: true }))
        // 认结构化信号（见 kbUnwrapToolResult）：工具**返回**失败信封时 registry 不置 isError，
        // 旧的"parse 散文"写法会把失败吞成成功 —— 那正是"经标注报成功、直连报失败"的根源。
        const u = kbUnwrapToolResult(r)
        if (u.failed) return { error: u.error }
        return { ok: true, output: (u.env && u.env.output) || u.env }
      },
    }),
  })

  reg(ctx, {
    name: 'kb_verify',
    description: '【kb · 只读】核账：annotations.jsonl 可解析、每条标注的源文件仍在且指纹吻合、影子方案与磁盘一致；问题清单随响应返回（clean 标志）。office_properties:true 时增一路——逐个只读打开 Office 标注文件读回 kb_* 自定义属性与登记比对（property-mismatch / property-missing，绝不猜）。',
    parameters: {
      root: { type: 'string', required: true, description: '已扫描的目录' },
      office_properties: { type: 'boolean', description: '读回 Office 文件的 kb_* 属性与登记比对（缺省关——每文件一次 COM 开销）' },
    },
    execute: async (args) => kbVerifyImpl(args, {
      // W2 读回闭环：office_properties:true 时经写回 blob 的只读动作（read_only_action）
      // 逐文件读 kb_* 属性——同一 COM 代码路径，读侧绝不落笔
      officeReader: async (absPath) => {
        const r = await runPython(KBW_CODE, JSON.stringify({
          path: absPath, properties: {}, read_only_action: true,
        }))
        return r.success ? r.output : { error: r.error }
      },
    }),
  })

  reg(ctx, {
    name: 'kb_write_back_office',
    description: '【kb · 变更】对单个 Office 文件写/删 kb_* 自定义文档属性——**只碰属性面板，绝不碰单元格/正文/幻灯片**。写后自动关掉重开核账（属性读回一致 + 内容快照 sha 一致）。preview 只读打开报告现值与将写值（物理不落笔）；managed 必须 confirm:true 双闸；rollback:true 删掉同名属性并同样核账。一次只动一个文件。⚠️ 口径：`write_back: true` 表示**写回操作成功**，**不等于文件字节变了**——若写入的值与原值相同，Office 判定无改动、不会重写盘，文件 mtime 不变（这不是失败）。落盘与否见 `output.disk_changed`。',
    parameters: {
      path: { type: 'string', required: true, description: 'Office 文件绝对路径（xlsx/xlsm/xls、docx/doc、pptx/ppt）' },
      properties: { type: 'object', required: true, description: '{kb_note/kb_taxonomy/kb_sensitivity/kb_fingerprint: 值}——键必须 kb_ 前缀；rollback 时键=要删的属性名、值忽略；值一律字符串，超 1024 字符截断并记 truncated' },
      rollback: { type: 'boolean', description: 'true = 删除 properties 里列出的 kb_* 属性（回滚=删属性即还原）' },
      mode: MODE_PARAMS.mode,
      confirm: MODE_PARAMS.confirm,
      overwrite: MODE_PARAMS.overwrite,
    },
    execute: async (args, ctlArgs) => {
      // 双闸（红线 3）：preview 之外，managed 必须 confirm:true，缺一不写。JS 侧先挡，
      // blob 内再查一遍 _ctl.confirm——两层各挡一次，漂移不了。
      const ctl = resolveCtl('kb_write_back_office', args)
      if (!ctl.valid) return { ok: false, error: '[MODE_INVALID] mode 须为 preview/managed/attached' }
      const preview = ctl.preview
      const props = args.properties || {}
      if (!args.path) return { ok: false, error: '[MISSING_PARAM] 缺少 path' }
      if (!Object.keys(props).length) return { ok: false, error: '[MISSING_PARAM] 缺少 properties（至少一个 kb_* 属性）' }
      if (!preview && !args.rollback && !ctl.confirm) {
        return { ok: false, error: '[RISKY_OP_NOT_CONFIRMED] 写回是变更动作：mode=managed 且 confirm:true 才落笔（缺一不写）' }
      }
      let pathReal
      try { pathReal = realpathSync.native(args.path) } catch {
        return { ok: false, error: '[OPEN_FAILED] 文件不存在: ' + args.path }
      }
      const guard = kbVaultGuard(pathReal)
      if (guard.guarded) {
        return { ok: false, error: `[VAULT_FORBIDDEN] 目标在 Obsidian vault（${guard.vault}）内——kb 工具不碰 vault，这是红线` }
      }
      // B4 防假成功保险丝（DSH 复跑 2026-10-06）：内侧核账是"关掉重开只读"读回来的，
      // 而 _shut 的 Quit() 被吞掉 —— 那个实例若没真退，重开可能读到**内存**而非磁盘，
      // 于是"属性读回一致"成立、磁盘却没变（实测 docx/pptx：报 write_back:true，而文件
      // 逐字节未动、包里没有 docProps/custom.xml）。
      // 这里加一条**不依赖根因**的外部判据：**新增**的属性必然改变文件字节（包要多出
      // custom.xml）。若「写前没有、写后声称有」而 size+mtime 一丝未动 ⇒ 保存没落盘。
      const st0 = statSync(pathReal, { bigint: true })
      let r
      try {
        r = await runPython(KBW_CODE, JSON.stringify({
          path: pathReal,
          properties: props,
          rollback: !!args.rollback,
          _ctl: ctl,
        }))
      } finally {
        // B5（DSH 验收 2026-10-06）：写回会 DispatchEx 起一个 Office 实例；blob 里 _shut 的
        // Quit() 送不走时它会一直活着、并**占着目标文件**（实测留下一个 WINWORD.EXE，
        // 后续读该文件报"正由另一进程使用"）。这里复用既有 pid 归属回收 ——
        // 按启动基线判"是不是我们起的"、**可见实例一律放过**（有人在看）、僵尸对象照收。
        // 不新发明一套：那套已经被 leakcheck 钉过。收尾失败不影响主流程。
        try { cleanupOwnedApps() } catch { /* 收尾尽力而为 */ }
      }
      if (!r.success) return { ok: false, error: r.error }
      const out = r.output
      if (preview) return { ok: true, preview: true, output: out }
      if (out && (out.action === 'write' || out.action === 'rollback')) {
        const st1 = statSync(pathReal, { bigint: true })
        const f = kbFuseTrips(out.action, props, out.properties_before || {}, out.applied, st0, st1)
        if (f.trips) {
          return { ok: false, error:
            `[KB_WRITE_FAILED] 保存未落盘：${f.byteChanging.join(',')} 属于**必然改字节**的改动，` +
            '但文件 size/mtime 一丝未动 —— 内侧核账很可能被骗过，**不要采信它的"读回一致"**' }
        }
        // 落盘与否如实回报，别让调用方从"属性一致"外推字节（write_back:true ≠ 文件变了）
        out.disk_changed = !f.unchanged
        out.disk_byte_changing = f.byteChanging
      }
      return { ok: true, output: out }
    },
  })

  reg(ctx, {
    name: 'kb_scan_root',
    description: '【kb · 只读】陌生目录的自建分类仓库第一步：登记一个目录树里的全部原生资料（扩展名+magic 双判据，认不出的一律进「不分类」桶并显式给原因码，绝不猜）。产物（registry/unclassified 两份 JSON）落插件状态目录 kb/<root_id>/，**绝不写进被测目录**；本机 Obsidian vault 及其子目录直接拒绝（VAULT_FORBIDDEN）。max_deep>0 时按配额只读开簿深扫 Excel 结构（表头/UsedRange/外链/超链接），并如实上报 VBA 通道可用性（信任中心未开 = VBA_ACCESS_DENIED，不假装可用）。',
    parameters: {
      root: { type: 'string', required: true, description: '要登记的目录（绝对路径）。Obsidian vault 及其子目录会被拒绝' },
      max_files: { type: 'number', description: '登记文件数上限（缺省 20000，防失控树）' },
      max_deep: { type: 'number', description: '深扫配额：只读开簿读取 Excel 结构的最大簿数（缺省 0 = 只做 FSO 枚举，不开 Excel）' },
      deep_exts: { type: 'array', description: '深扫対象扩展名（缺省 [".xlsx",".xlsm",".xls"]）' },
    },
    execute: async (args, ctl) => {
      // 守卫先于一切（降级态也生效）：VAULT_FORBIDDEN 不进 COM。
      // 返回**原始** {ok, error} 形态——reg() 的 finalize 负责定稿信封，这里不能再包一层。
      let rootReal
      try {
        rootReal = realpathSync.native(args.root)
      } catch {
        return { ok: false, error: '[OPEN_FAILED] 目录不存在: ' + args.root }
      }
      const guard = kbVaultGuard(rootReal)
      if (guard.guarded) {
        return {
          ok: false,
          error: `[VAULT_FORBIDDEN] 目标 ${rootReal} 与 Obsidian vault（${guard.vault}）重叠——kb 工具不碰 Obsidian，这是红线`,
        }
      }
      const rootId = createHash('sha256').update(rootReal).digest('hex').slice(0, 16)
      const r = await runPython(KB_CODE, JSON.stringify({
        root: rootReal,
        max_files: args.max_files,
        max_deep: args.max_deep,
        deep_exts: args.deep_exts,
        vocab: KB_POOL_EXTENSIONS,
        not_material: KB_NOT_MATERIAL,
        magic: KB_MAGIC,
      }))
      if (!r.success) return { ok: false, error: r.error }
      const out = r.output
      const dir = join(STATE_DIR, 'kb', rootId)
      mkdirSync(dir, { recursive: true })
      const registryPath = join(dir, 'registry.json')
      const unclassifiedPath = join(dir, 'unclassified.json')
      writeFileSync(registryPath, JSON.stringify(out, null, 1), 'utf-8')
      // P1.5-5：不分类报告是交付物——unrecognized 在前（要人眼看）、not-material 在后（登记即可）
      const byBucket = (a, b) => (a.reason || '').localeCompare(b.reason || '') ||
        (a.ext || a.format || '').localeCompare(b.ext || b.format || '')
      const unclassified = {
        unrecognized: [...(out.unclassified?.unrecognized || [])].sort(byBucket),
        not_material: [...(out.unclassified?.not_material || [])].sort(byBucket),
      }
      writeFileSync(unclassifiedPath, JSON.stringify(unclassified, null, 1), 'utf-8')
      out.artifacts = { dir, registry: registryPath, unclassified: unclassifiedPath }
      out.root_id = rootId
      return { ok: true, output: out }
    },
  })

  // ── kb_classify / kb_scheme_propose（P2-b）：机械分类 + 影子方案 ─────────
  // 冻结语义（红线 9）：not-material / 带 reason 的条目不参与分类与方案。
  // 分类只读；方案只落状态目录 scheme.json——影子目录，v1 永不 apply。

  const kbLoadRegistry = (rootReal) => {
    const rootId = createHash('sha256').update(rootReal).digest('hex').slice(0, 16)
    const p = join(STATE_DIR, 'kb', rootId, 'registry.json')
    if (!existsSync(p)) return { path: p, registry: null }
    try {
      return { path: p, registry: JSON.parse(readFileSync(p, 'utf-8')) }
    } catch (e) {
      return { path: p, registry: null, error: String(e) }
    }
  }

  const kbFrozen = (e) => e.get?.('pool') === 'not-material' || !!(e.get && e.get('reason')) ||
    (typeof e === 'object' && (e.pool === 'not-material' || !!e.reason))

  const kbMatchRule = (rule, entry) => {
    const pat = String(rule.pattern || '').toLowerCase()
    if (rule.kind === 'ext') return ('.' + (entry.format || '')) === pat
    if (rule.kind === 'name-contains') return (entry.title || entry.source_uri || '').toLowerCase().includes(pat)
    if (rule.kind === 'path-contains') return (entry.source_uri || '').toLowerCase().includes(pat)
    return false
  }

  const kbClassifyEntries = (registry, tax) => {
    const assigned = []
    const unassigned = []
    const conflicts = []
    for (const e of Object.values(registry.entries || {})) {
      if (kbFrozen(e)) continue                    // 红线 9：冻结桶不进分类流
      const hits = []
      for (const c of tax.categories) {
        for (const r of c.rules || []) {
          if (kbMatchRule(r, e)) { hits.push(c.id); break }
        }
      }
      const hitsU = [...new Set(hits)]
      if (hitsU.length === 1) {
        assigned.push({ source_uri: e.source_uri, fingerprint: e.fingerprint, taxonomy_id: hitsU[0] })
      } else if (hitsU.length === 0) {
        unassigned.push({ source_uri: e.source_uri, fingerprint: e.fingerprint })
      } else {
        conflicts.push({ source_uri: e.source_uri, fingerprint: e.fingerprint, taxonomy_ids: hitsU })
      }
    }
    return { assigned, unassigned, conflicts }
  }

  /** 解析 knowlp_search 在注册表里的**实际名字**。
   *  §1.4 修正（DSH 验收 2026-10-06）：MCP 工具经 dsh-mcp-client 注册时用的是 public name
   *  `mcp__<serverName>__<rawName>`（其 publicToolName()），**不保留 raw 别名**；而
   *  dsh-tools 的 `get(name, scope)` 是**精确查表**。所以按 `'knowlp_search'` 直查**恒 undefined**
   *  —— 那**不是"knowlp 不在"**，是我们查错了名字。旧实现据此掉到 unavailable，
   *  于是 P3-b **从未真正跑通过**。
   *  三级解析：① 原名直查（原生插件 / 非 MCP 宿主）② `mcp__<server>__<raw>` ③ 扫 `schemas()`。 */
  const kbResolveKnowlp = (scope) => {
    const t = _ctx?.tools
    if (!t || typeof t.get !== 'function') return null
    const atom = 'knowlp_search'
    const tryName = (n) => { const d = t.get(n, scope); return d ? { name: n, def: d } : null }
    const direct = tryName(atom)
    if (direct) return { ...direct, how: 'raw' }
    for (const server of ['knowlp']) {
      const hit = tryName(`mcp__${server}__${atom}`)
      if (hit) return { ...hit, how: 'mcp-public' }
    }
    if (typeof t.schemas === 'function') {
      try {
        const found = (t.schemas(scope) || []).find(
          (s) => typeof (s && s.name) === 'string' && s.name.endsWith(`__${atom}`))
        if (found) {
          const d = t.get(found.name, scope)
          if (d) return { name: found.name, def: d, how: 'mcp-scan' }
        }
      } catch { /* 扫描失败 = 当作没解析到，走 unavailable，不猜 */ }
    }
    return null
  }

  /** P3-b 语义缝降级梯子（工单 14）：
   *  approved = ctx.tools.execute（走审批）；direct = 拿到工具对象直接 .execute（绕过审批，仅降级）；
   *  unavailable = 解析不到 knowlp。断链时 kb_classify 显式声明 semantics: unavailable。
   *  返回里带 `how`（名字是怎么解析到的）——"查不到"与"服务不可用"必须在读数上分得开。 */
  const kbSemanticsLadder = (scope) => {
    const hit = kbResolveKnowlp(scope)
    if (hit && typeof _ctx?.tools?.execute === 'function') {
      return { step: 'approved', name: hit.name, how: hit.how }
    }
    if (hit && hit.def && typeof hit.def.execute === 'function') {
      return { step: 'direct', tool: hit.def, name: hit.name, how: hit.how }
    }
    return { step: 'unavailable', how: 'unresolved' }
  }

  const kbSemanticAssign = async (unassigned, tax, cap = 10, exec = null) => {
    const seam = kbSemanticsLadder(exec?.agent)
    if (seam.step === 'unavailable') {
      // 降级梯子②：knowlp 不可用 → 返回「下一步调用建议」，由 agent 自己调（不假装做了语义判断）
      const first = unassigned[0]
      const q = first ? decodeURIComponent((first.source_uri || '').split('/').pop() || '') : ''
      return { semantics: 'unavailable', seam_how: seam.how, semantic_assignments: [],
               suggestion: first ? { tool: 'knowlp_search', arguments: { query: q, limit: 5 } } : null }
    }
    const labels = tax.categories.map((c) => ({ id: c.id, label: c.label }))
    const semAssign = []
    const probe = []
    const seamErrors = []
    let degraded = seam.step === 'direct'
    for (const u of unassigned.slice(0, cap)) {
      const title = decodeURIComponent((u.source_uri || '').split('/').pop() || '')
      const call = { tool: 'knowlp_search', query: title, arguments: { query: title, limit: 5 } }
      probe.push({ tool: call.tool, arguments: call.arguments })
      let hits = []
      try {
        let r
        if (seam.step === 'approved') {
          // ★10 确证（dsh-tools lib/index.js:3116 + :1298-1308）：registry.execute 收**单对象**
          // {callId, name, arguments}；返回 {isError, content:[...]}（MCP 内容块形态）——
          // 旧的 ('knowlp_search', args) 双位置参数形状是错的，异常曾被静默吞成 0 命中
          // 用**解析到的实际名字**派发（§1.4）——写死 'knowlp_search' 在 MCP 接入下必然查不到
          r = await _ctx.tools.execute(kbNestedCall(exec, seam.name || 'knowlp_search', call.arguments))
        } else {
          r = await seam.tool.execute(call.arguments)
        }
        // 认结构化信号（见 kbUnwrapToolResult）：**不再**对渲染散文做 JSON.parse
        const u = kbUnwrapToolResult(r)
        if (u.failed) throw new Error(u.error)
        const env = u.env || {}
        hits = (((env.output && env.output.hits) || env.hits) || [])
          .map((h) => String(h.title || h.name || ''))
      } catch (e) {
        // ★8：绝不静默——seam 失败降级为 approved-failed，错误进 output
        seamErrors.push({ query: call.query, error: String(e) })
        continue
      }
      for (const c of labels) {
        if (hits.some((h) => h.includes(c.label))) {
          semAssign.push({ source_uri: u.source_uri, fingerprint: u.fingerprint,
                           taxonomy_id: c.id, reason: 'semantic' })
          break
        }
      }
    }
    let semantics = degraded ? 'direct-degraded' : 'approved'
    if (seamErrors.length) semantics = seam.step === 'approved' ? 'approved-failed' : 'direct-failed'
    return { semantics, seam_how: seam.how, semantic_assignments: semAssign, suggestion: null, probe,
             seam_errors: seamErrors }
  }

  const kbClassifyTool = async (args, exec) => {
    let rootReal
    try {
      rootReal = realpathSync.native(args.root)
    } catch {
      return { ok: false, error: '[OPEN_FAILED] 目录不存在: ' + args.root }
    }
    const { path: regPath, registry } = kbLoadRegistry(rootReal)
    if (!registry) {
      return { ok: false, error: '[OPEN_FAILED] 该目录还没有 registry——先跑 kb_scan_root' }
    }
    const tax = kbLoadTaxonomy(kbTaxonomyPath(rootReal))
    if (!tax.categories.length) {
      return { ok: false, error: '[NO_TAXONOMY] 分类学为空——先用 kb_taxonomy_set 建类别' }
    }
    const result = kbClassifyEntries(registry, tax)
    // §1.8-2 修正（DSH 验收 2026-10-06）：`unavailable` 曾承载三种互不相同的处境 ——
    // knowlp 不可用 / 没传 use_semantics / 无待归类项 —— 读数同形，验收时无法区分。
    // 现三态分开：not-requested（没要）→ not-needed（要了但没东西可判）→ unavailable（要了但连不上）。
    let sem = { semantics: 'not-requested', semantic_assignments: [], suggestion: null, seam_how: 'n/a' }
    if (args.use_semantics === true && !result.unassigned.length) {
      sem = { semantics: 'not-needed', semantic_assignments: [], suggestion: null, seam_how: 'n/a' }
    }
    if (args.use_semantics === true && result.unassigned.length) {
      sem = await kbSemanticAssign(result.unassigned, tax, 10, exec)
      for (const sa of sem.semantic_assignments) {
        result.assigned.push(sa)
        result.unassigned = result.unassigned.filter((u) => u.source_uri !== sa.source_uri)
      }
    }
    return {
      ok: true,
      output: { root: rootReal, read_only: true, semantics: sem.semantics,
                seam_how: sem.seam_how || 'n/a',
                semantic_assignments: sem.semantic_assignments,
                seam_errors: sem.seam_errors || [],
                next_call_suggestion: sem.suggestion,
                assigned: result.assigned, unassigned: result.unassigned,
                conflicts: result.conflicts,
                stats: { assigned: result.assigned.length, unassigned: result.unassigned.length,
                         conflicts: result.conflicts.length } },
    }
  }

  const kbSchemePropose = (args) => {
    let rootReal
    try {
      rootReal = realpathSync.native(args.root)
    } catch {
      return { ok: false, error: '[OPEN_FAILED] 目录不存在: ' + args.root }
    }
    const { registry } = kbLoadRegistry(rootReal)
    if (!registry) {
      return { ok: false, error: '[OPEN_FAILED] 该目录还没有 registry——先跑 kb_scan_root' }
    }
    const tax = kbLoadTaxonomy(kbTaxonomyPath(rootReal))
    if (!tax.categories.length) {
      return { ok: false, error: '[NO_TAXONOMY] 分类学为空——先用 kb_taxonomy_set 建类别' }
    }
    const { assigned, unassigned, conflicts } = kbClassifyEntries(registry, tax)
    const byFp = new Map(Object.values(registry.entries || {}).map((e) => [e.fingerprint, e]))
    const usedNames = new Set()
    const moves = []
    for (const a of assigned) {
      const e = byFp.get(a.fingerprint)
      if (!e) continue
      const cat = tax.categories.find((c) => c.id === a.taxonomy_id)
      const segs = []
      let cur = cat
      const chain = []
      while (cur) { chain.unshift(cur.label); cur = cur.parent_id ? tax.categories.find((c) => c.id === cur.parent_id) : null }
      segs.push(...chain)                                     // 深度 = parent 链
      let name = e.source_uri.split('/').pop() || 'unnamed'
      if (usedNames.has(segs.join('/') + '/' + name)) {
        name = a.fingerprint.slice(0, 8) + '-' + name          // 命名策略：重名加指纹前缀
      }
      usedNames.add(segs.join('/') + '/' + name)
      segs.push(name)
      moves.push({
        fingerprint: a.fingerprint,
        from: e.source_uri,
        to: 'file://' + segs.join('/'),
        taxonomy_id: a.taxonomy_id,
        reason: 'taxonomy-rule',
      })
    }
    const plan_id = createHash('sha256')
      .update(rootReal + '|' + tax.version + '|' + new Date().toISOString())
      .digest('hex').slice(0, 12)
    const scheme = {
      plan_id, root: rootReal, generated_at: new Date().toISOString(),
      taxonomy_version: tax.version,
      naming: 'label-chain/原文件名，重名加指纹前 8 位',
      moves,
      untouched: { unassigned: unassigned.length, conflicts: conflicts.length,
                   frozen: Object.values(registry.entries || {}).filter((e) => kbFrozen(e)).length },
    }
    const rootId = createHash('sha256').update(rootReal).digest('hex').slice(0, 16)
    const dir = join(STATE_DIR, 'kb', rootId)
    mkdirSync(dir, { recursive: true })
    const schemePath = join(dir, 'scheme-' + plan_id + '.json')
    writeFileSync(schemePath, JSON.stringify(scheme, null, 1), 'utf-8')
    return {
      ok: true,
      output: { root: rootReal, plan_id, scheme_path: schemePath,
                moves: moves.length, untouched: scheme.untouched,
                read_only: true, note: '影子方案——只出方案不碰文件；apply 属后续版本' },
    }
  }

  reg(ctx, {
    name: 'kb_taxonomy_get',
    description: '【kb · 只读】读取某目录的分类学（agent 自长分类，非固定枚举），并附漂移审计（重叠分类/环/悬空指针）。',
    parameters: {
      root: { type: 'string', required: true, description: '分类学所属的目录（与 kb_scan_root 同一 root 即同一份分类学）' },
    },
    execute: async (args) => kbTaxonomyGet(args),
  })

  reg(ctx, {
    name: 'kb_taxonomy_set',
    description: '【kb · 变更】演化分类学：add / rename / merge / delete。merge 会把被并类别的规则并入目标并记 supersedes（沿用 evidence 语义）；delete 遇有子类别需 force。每次 set 自带漂移审计结果（重叠分类=规则冲突信号）。',
    parameters: {
      root: { type: 'string', required: true, description: '分类学所属的目录' },
      ops: {
        type: 'array', required: true,
        description: '操作数组：{op:"add", category:{id,label,parent_id?,description?,rules?}} / {op:"rename", id, label?, description?} / {op:"merge", into_id, from_ids:[...]} / {op:"delete", id, force?}',
      },
    },
    execute: async (args) => kbTaxonomySet(args),
  })

  console.log(`[dsh-office-com] plugin loaded${degraded ? ' (degraded)' : ''}`)
}
