// faults.mjs — 异常与故障注入（工单 §五 点名必须覆盖的那几项）
//   保存失败（只读文件）/ 文件被占用 / 宏执行失败 / OfficeMCP 断线重连 / 值类型边界
//
// 这一份专测「只在真机上才暴露」的路径：信封在这些场景里最容易说谎——
// 报成 ok、报错码不稳、或声称已保存但其实没落盘。
//
// 运行: node test/faults.mjs   （需本机 Office + officemcp）
import { apply, runPython } from '../lib/index.mjs'
import { existsSync, rmSync, statSync, chmodSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir, homedir } from 'node:os'
import { spawnSync } from 'node:child_process'

const tools = {}
apply({ tools: { register: (t) => { tools[t.name] = t } } })

let fail = 0
function check(name, cond, detail) {
  console.log(`[${cond ? 'PASS' : 'FAIL'}] ${name}${detail ? ' — ' + detail : ''}`)
  if (!cond) fail++
}

async function raw(lines, argsObj) {
  const code = ['import json, pythoncom, os', 'pythoncom.CoInitialize()',
    'a = json.loads(data) if data else {}',
    "if a.get('path'): a['path'] = a['path'].replace('/', os.sep)", ...lines].join('\n')
  const r = await runPython(code, JSON.stringify(argsObj || {}))
  if (!r.success) throw new Error(r.error)
  return r.output
}

const D = tmpdir()
const F = (n) => join(D, n)
const P = (p) => p.replace(/\\/g, '/')
const reset = (p) => { if (existsSync(p)) rmSync(p) }

/** 只读文件（模拟"文件被占用/不可写"这一整类）。收尾务必恢复，否则后续用例全线崩。 */
function setReadOnly(p, ro) {
  try { chmodSync(p, ro ? 0o444 : 0o666) } catch { /* 忽略 */ }
}

console.log('══ 异常与故障注入 ══\n')

// ══ A. 保存失败：只读文件 ═══════════════════════════════════════
{
  const B = F('f-ro.xlsx')
  reset(B)
  await raw(['excel = Officer.Excel', 'excel.DisplayAlerts = False', 'excel.Visible = False',
    'wb = excel.Workbooks.Add()', "wb.Sheets(1).Range('A1').Value = '原值'",
    "wb.SaveAs(a['path'])", 'wb.Close(False)', "output = json.dumps({'ok': True})"], { path: P(B) })
  setReadOnly(B, true)
  const m0 = statSync(B).mtimeMs

  const r = await tools.excel_write_range.execute({ path: B, range: 'B1', value: '想写进去' })
  check('保存失败：信封报 ok=false 且码为 SAVE_FAILED', !r.ok && r.error_code === 'SAVE_FAILED', `[${r.error_code}] ${String(r.error).slice(0, 90)}`)
  check('保存失败：可重试（多为瞬时占用）', r.retryable === true, `retryable=${r.retryable}`)
  check('保存失败：saved 不得虚报为 true', r.saved !== true, `saved=${r.saved}`)
  check('保存失败：managed 模式回滚，磁盘文件未被改动', statSync(B).mtimeMs === m0, `${m0} vs ${statSync(B).mtimeMs}`)

  setReadOnly(B, false)
  const r2 = await tools.excel_read_range.execute({ path: B, range: 'A1' })
  check('保存失败：文件仍可正常读取、原值未损', r2.ok && r2.output?.value === '原值', JSON.stringify(r2.output?.value))
}

// ══ B. 文件被占用：附件模式下的部分修改必须被标识出来 ═════════════
{
  const B = F('f-partial.xlsx')
  reset(B)
  await raw(['excel = Officer.Excel', 'excel.DisplayAlerts = False', 'excel.Visible = False',
    'wb = excel.Workbooks.Add()', 'wb.SaveAs(a["path"])', 'wb.Close(False)', "output = json.dumps({'ok': True})"], { path: P(B) })
  // 附着到已打开的工作簿（不带 path），制造"改了内存但没落盘"的场景
  await raw(['excel = Officer.Excel', 'excel.DisplayAlerts = False',
    'excel.Workbooks.Open(a["path"])', "output = json.dumps({'ok': True})"], { path: P(B) })
  const r = await tools.excel_write_range.execute({ range: 'A1', value: 'x' })
  check('附着模式：改了但 saved=false（信封不虚报落盘）', r.ok && r.changed === true && r.saved === false, JSON.stringify({ c: r.changed, s: r.saved }))
  await raw(['excel = Officer.Excel', 'excel.DisplayAlerts = False',
    "wb = [w for w in excel.Workbooks if w.Name == a['name']]", 'if wb: wb[0].Close(False)', "output = json.dumps({'ok': True})"], { name: 'f-partial.xlsx' })
}

// ══ C. 宏执行失败 ═══════════════════════════════════════════════
{
  await tools.excel_new.execute()
  let r = await tools.excel_vba_run.execute({ macro: '__no_such_macro__' })
  check('宏：不给 confirm 时被高危门控挡住', !r.ok && r.error_code === 'RISKY_OP_NOT_CONFIRMED', `[${r.error_code}]`)

  r = await tools.excel_vba_run.execute({ macro: '__no_such_macro__', mode: 'preview' })
  check('宏：preview 不被门控挡、且点明确认位未就位', r.ok && r.preview === true && r.output?.requires_confirm === true && r.output?.confirm_present === false,
    r.ok ? JSON.stringify({ need: r.output?.requires_confirm, has: r.output?.confirm_present }) : `[${r.error_code}]`)

  r = await tools.excel_vba_run.execute({ macro: '__no_such_macro__', confirm: true })
  check('宏：执行失败报 VBA_FAILED（不是 UNKNOWN，COM 文本随语言变不能靠模糊匹配）', !r.ok && r.error_code === 'VBA_FAILED', `[${r.error_code}] ${String(r.error).slice(0, 80)}`)
}

// ══ D. 值类型边界：空 / 日期 / 金额 / 异常类型 ═══════════════════
{
  await tools.excel_new.execute()

  // 空值写入：显式 None 与空串都不该炸
  let r = await tools.excel_write_range.execute({ range: 'A1', value: null })
  check('类型：写 null 不炸', r.ok, `[${r.error_code}] ${String(r.error).slice(0, 80)}`)
  r = await tools.excel_write_range.execute({ range: 'A2', value: '' })
  check('类型：写空串不炸', r.ok, `[${r.error_code}]`)

  // 数字 / 布尔 / 负数的往返
  await tools.excel_write_range.execute({ range: 'B1:C1', value: [[-123.45, 0], [true, false]] })
  r = await tools.excel_read_range.execute({ range: 'B1:C2' })
  const v = JSON.stringify(r.output?.value)
  check('类型：负数/零/布尔往返一致', v.includes('-123.45') && v.includes('0'), v)

  // 日期：Excel 会把它存成真日期，读回来是 datetime → 插件必须归一成 ISO 字符串而不是崩
  await tools.excel_write_range.execute({ range: 'D1', value: '2026-08-03' })
  r = await tools.excel_read_range.execute({ range: 'D1' })
  const dv = String(r.output?.value)
  check('类型：日期往返不崩、且是 ISO 形态', r.ok && /2026-08-03/.test(dv), dv)

  // 公式结果里的异常类型：分母为 0 → 读回不应抛，而应是可辨识的形态
  await tools.excel_formula_set.execute({ range: 'E1', formula: '=1/0' })
  r = await tools.excel_read_range.execute({ range: 'E1' })
  check('类型：错误公式读回不抛异常（返回可辨识形态）', r.ok, r.ok ? String(r.output?.value) : `[${r.error_code}] ${String(r.error).slice(0, 80)}`)

  // 非法区域引用 → 必须是稳定码而不是 UNKNOWN
  r = await tools.excel_read_range.execute({ range: 'NOT_A_RANGE' })
  check('类型：非法区域报稳定码（非 UNKNOWN）', !r.ok && r.error_code !== 'UNKNOWN', `[${r.error_code}]`)
}

// ══ E. OfficeMCP 断线重连 ══════════════════════════════════════
{
  let r = await tools.office_apps.execute()
  check('断线前：通道可用', r.ok, r.ok ? '' : `[${r.error_code}]`)

  const lockPath = join(homedir(), '.dsh-office-com', 'office.lock')
  let killed = null
  try {
    const lock = JSON.parse(String(readFileSync(lockPath, 'utf8')))
    if (lock?.pid) {
      spawnSync('taskkill', ['/F', '/PID', String(lock.pid)], { encoding: 'utf8' })
      killed = lock.pid
    }
  } catch { /* 没锁文件就算了 */ }

  if (killed) {
    console.log(`[info] 已杀掉 SSE 服务 pid ${killed}，观察重连`)
    await new Promise((r2) => setTimeout(r2, 1500))
    const first = await tools.office_apps.execute()
    check('断线后第一次调用：要么自动重连成功、要么报可重试的通道错误',
      first.ok || (first.error_code === 'CHANNEL_UNAVAILABLE' && first.retryable === true),
      first.ok ? '自动重连成功' : `[${first.error_code}] retryable=${first.retryable}`)
    const second = await tools.office_apps.execute()
    check('断线后第二次调用必须成功（证明重连真的生效）', second.ok, second.ok ? '' : `[${second.error_code}] ${String(second.error).slice(0, 80)}`)
  } else {
    console.log('[skip] 读不到锁文件 pid，跳过断线重连用例')
  }
}

// ══ 收尾：本测试自己要收干净 ════════════════════════════════════
// 测试里用 excel_new 造的几个工作簿会**按设计**留在打开状态（那正是 excel_new 的行为，
// 由 leakcheck 用例 C 守着），所以这里主动关掉再断言归零——顺便也验了批量关闭有效。
{
  const left = await raw(['excel = Officer.Excel', 'excel.DisplayAlerts = False',
    '[w.Close(False) for w in list(excel.Workbooks)]',
    'output = json.dumps([w.Name for w in excel.Workbooks])'], {})
  check('收尾：关闭后无遗留打开的工作簿', Array.isArray(left) && left.length === 0, JSON.stringify(left))
}

console.log(fail === 0 ? '\n[FAULTS] ALL PASS' : `\n[FAULTS] ${fail} FAILURE(S)`)
console.log('__DONE__')
setTimeout(() => process.exit(fail === 0 ? 0 : 1), 500)
