// 进程泄漏检查：宿主进程退出后，插件起的 Office 进程是否变成孤儿；以及可见实例是否被误杀。
// 运行: node test/leakcheck.mjs   （需本机 Office + officemcp）
//
// 背景：officemcp 的 Officer.Excel 是 GetActiveObject-else-Dispatch 并**永久缓存**的 COM 引用。
// 我们只 kill python SSE 子进程（TerminateProcess，atexit 不跑），Excel 只要还开着工作簿就不会
// 自己退 → 宿主退出后留下不可见的孤儿 EXCEL.EXE。lib 的退出收尾只回收「启动前不存在 且 不可见」的实例。
//
// 三个用例：
//   A excel_open/write/read（managed 模式，Excel 全程不可见）→ 宿主退出后不该有残留
//   B office_launch(visible:true)  → 可见实例必须存活（收尾不能误杀用户看得见的东西）
//   C excel_new                    → 必须置可见并留给用户（不置可见就等于凭空多一个够不着的孤儿）
import { spawnSync, spawn } from 'node:child_process'
import { existsSync, unlinkSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { findPython } from '../lib/index.mjs'

const TMP_BOOK = join(tmpdir(), 'dsh-leak-check.xlsx')
const LIB_URL = new URL('../lib/index.mjs', import.meta.url).href

function pidsOf(image) {
  const r = spawnSync('tasklist', ['/FI', `IMAGENAME eq ${image}`, '/FO', 'CSV', '/NH'], { encoding: 'utf8' })
  return (r.stdout || '').split('\n').map((l) => l.match(/^"[^"]+","(\d+)"/)).filter(Boolean).map((m) => Number(m[1]))
}
const excelPids = () => pidsOf('EXCEL.EXE')
const killPid = (pid) => spawnSync('taskkill', ['/F', '/PID', String(pid)], { encoding: 'utf8' })
const sleep = (ms) => new Promise((r) => setTimeout(r, ms))

/** 在子进程里跑一段真实插件链路；host 退出后的效应只有跨进程才观察得到。 */
function runChild(body) {
  return new Promise((resolve) => {
    const scenario = `
import { apply } from '${LIB_URL}'
const tools = {}
apply({ tools: { register: (t) => { tools[t.name] = t } } })
const log = (n, r) => console.log(n, r.ok ? 'ok' : \`FAIL[\${r.error_code}] \${r.error}\`)
${body}
console.log('__DONE__')
setTimeout(() => process.exit(0), 300) // SSE 长连接占着事件循环，不显式退出子进程永远不会结束
`
    const child = spawn(process.execPath, ['--input-type=module', '-e', scenario], { stdio: ['ignore', 'pipe', 'pipe'] })
    let out = ''
    let err = ''
    child.stdout.on('data', (d) => { out += d })
    child.stderr.on('data', (d) => { err += d })
    const guard = setTimeout(() => { child.kill(); err += '\n[leakcheck] 子进程 150s 未退出，已强杀' }, 150000)
    child.on('exit', (code) => { clearTimeout(guard); resolve({ code, out, err }) })
  })
}

function quitViaCom(py, app) {
  const code = [
    'import pythoncom, win32com.client',
    'pythoncom.CoInitialize()',
    `app = win32com.client.GetActiveObject('${app}.Application')`,
    'app.DisplayAlerts = False',
    'app.Quit()',
  ].join(';')
  return spawnSync(py, ['-c', code], { encoding: 'utf8', timeout: 20000 })
}

/** 造一个真实 xlsx（写工具要 Open 已存在的文件），并把准备用的 Excel 收干净——否则它的残留进程会被
 *  当成「preexisting」记进锁文件，让用例 A 变成假通过。 */
function prepareWorkbook(py) {
  spawnSync(py, ['-c', [
    'import pythoncom, win32com.client, os',
    'pythoncom.CoInitialize()',
    `p = r'${TMP_BOOK}'`,
    'if os.path.exists(p):',
    '    os.remove(p)',
    'excel = win32com.client.Dispatch("Excel.Application")',
    'excel.DisplayAlerts = False',
    'wb = excel.Workbooks.Add()',
    'wb.SaveAs(p)',
    'wb.Close(False)',
    'excel.Quit()',
  ].join('\n')], { encoding: 'utf8', timeout: 60000 })
  if (!existsSync(TMP_BOOK)) return false
  quitViaCom(py, 'Excel') // 脚本里的 Quit 是异步的，进程可能比它先退，补一枪
  return true
}

async function main() {
  const py = findPython()
  if (!py) {
    console.error('[leakcheck] FAIL: 找不到 officemcp python')
    process.exit(1)
  }
  console.log(`[leakcheck] python = ${py}`)

  if (excelPids().length) {
    console.error(`[leakcheck] FAIL: 检查前已有 Excel 在跑（pid ${excelPids().join(',')}），无法区分是不是我们起的——请先关掉 Excel`)
    process.exit(1)
  }
  if (!prepareWorkbook(py)) {
    console.error('[leakcheck] FAIL: 准备测试工作簿失败')
    process.exit(1)
  }
  for (const pid of excelPids()) killPid(pid) // 准备阶段的一律强杀，保证基线干净
  await sleep(1500)
  if (excelPids().length) {
    console.error(`[leakcheck] FAIL: 准备阶段 Excel 未清干净（pid ${excelPids().join(',')}）`)
    process.exit(1)
  }
  console.log('[leakcheck] 基线: EXCEL=0（干净）')

  // ── 用例 A：全不可见链路（managed 模式，Excel 只为干活而存在）→ 宿主退出后不该有残留 ──
  console.log('[leakcheck] A 不可见链路（excel_open → write → read）…')
  const a = await runChild(`
log('excel_open', await tools.excel_open.execute({ path: ${JSON.stringify(TMP_BOOK)} }))
log('excel_write_range', await tools.excel_write_range.execute({ path: ${JSON.stringify(TMP_BOOK)}, sheet: 'Sheet1', range: 'A1:B2', value: [[1,2],[3,4]] }))
log('excel_read_range', await tools.excel_read_range.execute({ path: ${JSON.stringify(TMP_BOOK)}, sheet: 'Sheet1', range: 'A1:B2' }))`)
  console.log(a.out.trim() || a.err.trim().slice(-400))
  await sleep(3000)
  const leakedA = excelPids()
  console.log(leakedA.length
    ? `[leakcheck] A LEAK: 宿主退出后残留 EXCEL pid ${leakedA.join(',')}`
    : '[leakcheck] A PASS: 宿主退出后无残留')
  for (const pid of leakedA) { quitViaCom(py, 'Excel'); killPid(pid) }
  await sleep(1500)

  // ── 用例 B：可见实例（=用户看得见的工作簿）→ 收尾必须放过 ──
  console.log('[leakcheck] B 可见实例（office_launch visible:true）…')
  const b = await runChild(`log('office_launch', await tools.office_launch.execute({ app: 'Excel', visible: true }))`)
  console.log(b.out.trim() || b.err.trim().slice(-400))
  await sleep(3000)
  const survivedB = excelPids()
  console.log(survivedB.length
    ? `[leakcheck] B PASS: 可见实例存活（pid ${survivedB.join(',')}），收尾没误杀`
    : '[leakcheck] B FAIL: 可见实例被误杀了')
  for (const pid of survivedB) killPid(pid)
  await sleep(1500)

  // ── 用例 C：excel_new 必须把 Excel 置可见（用户看得见、够得着的工作簿）→ 存活 ──
  console.log('[leakcheck] C excel_new 置可见…')
  const c = await runChild(`log('excel_new', await tools.excel_new.execute())`)
  console.log(c.out.trim() || c.err.trim().slice(-400))
  await sleep(3000)
  const survivedC = excelPids()
  console.log(survivedC.length
    ? `[leakcheck] C PASS: 新建工作簿的 Excel 存活（pid ${survivedC.join(',')}）`
    : '[leakcheck] C FAIL: excel_new 的实例被回收了（应置可见并留给用户）')
  for (const pid of survivedC) killPid(pid)

  try { unlinkSync(TMP_BOOK) } catch { /* 没建成 */ }
  const pass = leakedA.length === 0 && survivedB.length > 0 && survivedC.length > 0
  console.log(pass ? '[leakcheck] PASS' : '[leakcheck] FAIL')
  process.exit(pass ? 0 : 2)
}

main().catch((e) => {
  console.error('[leakcheck] EXCEPTION:', e.message)
  process.exit(1)
})
