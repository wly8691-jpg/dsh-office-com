// flagship.mjs — 任务回归测试（工单 §五 的最上层）：会计旗舰链路整条跑两遍，
// 验证「重复执行能得到正确结果、且不破坏工作簿」，并打印 §八 要求的结构化验证报告。
//
// 与其余测试的分工：
//   register.mjs  = 纯函数契约（CI，无需 Office）
//   headless.mjs  = 工具编排 / 输入校验 / 返回结构
//   flagship.mjs  = 一整条业务链的**任务级**正确性与可重复性（本文件）
//   leakcheck.mjs = 跨进程的 Office 进程残留
//
// 运行: node test/flagship.mjs   （需本机 Office，且跑前请先关掉 Excel）
import { apply, runPython } from '../lib/index.mjs'
import { existsSync, rmSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'

const tools = {}
apply({ tools: { register: (t) => { tools[t.name] = t } } })

// ── 辅助 ─────────────────────────────────────────────────────
const REPORT = []
let failures = 0
function check(metric, cond, detail) {
  const mark = cond ? 'PASS' : 'FAIL'
  console.log(`[${mark}] ${metric}${detail ? ' — ' + detail : ''}`)
  REPORT.push({ metric, ok: !!cond, detail: detail ?? null })
  if (!cond) failures++
  return cond
}

async function call(name, args) {
  const t = tools[name]
  if (!t) throw new Error(`tool not registered: ${name}`)
  const r = await t.execute(args || {})
  if (!r.ok) throw new Error(`${name} -> [${r.error_code}] ${r.error}`)
  return r
}

/** 裸 COM 检查：路径走 data 通道传，避开反斜杠转义。 */
async function raw(pyLines, argsObj) {
  const code = ['import json, pythoncom', 'pythoncom.CoInitialize()', 'a = json.loads(data) if data else {}',
    "if a.get('path'): a['path'] = a['path'].replace('/', chr(92))", ...pyLines].join('\n')
  const r = await runPython(code, JSON.stringify(argsObj || {}))
  if (!r.success) throw new Error(r.error)
  return r.output
}

const DIR = tmpdir()
const BOOK = join(DIR, 'flagship-accounting.xlsx')
const P = BOOK.replace(/\\/g, '/')
if (existsSync(BOOK)) rmSync(BOOK)

// ── 输入：一份脱敏交易数据（§八 第 1 项）──────────────────────
const TRANSACTIONS = [
  { date: '2026-08-03', desc: '采购办公用品', account: '办公费', debit: 1200, credit: 0 },
  { date: '2026-08-03', desc: '采购办公用品', account: '银行存款', debit: 0, credit: 1200 },
  { date: '2026-08-07', desc: '支付差旅费', account: '差旅费', debit: 860, credit: 0 },
  { date: '2026-08-07', desc: '支付差旅费', account: '银行存款', debit: 0, credit: 860 },
  { date: '2026-08-15', desc: '计提工资', account: '管理费用', debit: 5000, credit: 0 },
  { date: '2026-08-15', desc: '计提工资', account: '应付职工薪酬', debit: 0, credit: 5000 },
]
const JOURNAL_TOTAL = TRANSACTIONS.reduce((s, e) => s + e.debit, 0)

console.log('═══ 会计旗舰链路任务回归 ═══\n')

// ── 0. 造出「模板」文件（§八 第 2 项）：带表头的空白账簿 ────────
try {
  await raw([
    'excel = Officer.Excel', 'excel.DisplayAlerts = False', 'excel.Visible = False',
    'wb = excel.Workbooks.Add()',
    "ws = wb.Sheets(1)",
    "ws.Cells(1,1).Value = '日期'; ws.Cells(1,2).Value = '摘要'; ws.Cells(1,3).Value = '科目'; ws.Cells(1,4).Value = '借方'; ws.Cells(1,5).Value = '贷方'",
    "ws.Name = '日记账'",
    "wb.SaveAs(a['path'])", 'wb.Close(False)',
    "output = json.dumps({'created': True})",
  ], { path: P })
} catch (e) {
  console.error(`无法创建测试工作簿（本机 Office / OfficeMCP 是否就绪？）: ${e.message}`)
  process.exit(2)
}

// ── 1. 一次 preview（§八 第 4 项）：真开文件看，一个字节都不改 ──
const beforeMtime = await raw(["import os", "output = json.dumps(os.path.getmtime(a['path']))"], { path: P })
const pv = await call('excel_journal_post', { path: P, sheet: '日记账', entries: TRANSACTIONS, mode: 'preview' })
check('preview：借贷平衡与合计算对', pv.output.balanced === true && pv.output.debit_total === JOURNAL_TOTAL, JSON.stringify({ dr: pv.output.debit_total, cr: pv.output.credit_total }))
check('preview：目标行与来源属实', pv.output.target.start_row === 2 && pv.output.entries === TRANSACTIONS.length, JSON.stringify(pv.output.target))
const afterMtime = await raw(["import os", "output = json.dumps(os.path.getmtime(a['path']))"], { path: P })
check('preview：未改动文件（mtime 不变）', beforeMtime === afterMtime, `${beforeMtime} vs ${afterMtime}`)

// ── 2. 一次 managed 执行（§八 第 5 项）：打开-改-校验-保存-关闭 ──
let r = await call('excel_journal_post', { path: P, sheet: '日记账', entries: TRANSACTIONS })
check('managed 入账：落盘且非幂等跳过', r.changed === true && r.saved === true && r.output.idempotent_skip === false, JSON.stringify({ changed: r.changed, saved: r.saved, posted: r.output.posted }))
check('managed 入账：借贷平衡识别正确', r.output.balanced === true && r.output.debit_total === JOURNAL_TOTAL && r.output.credit_total === JOURNAL_TOTAL, JSON.stringify({ dr: r.output.debit_total, cr: r.output.credit_total }))

r = await call('excel_ledger_gen', { path: P, journal_sheet: '日记账' })
const ledgerSheet = r.output.output_sheet
// 日记账里出现 5 个不同科目 → 总账 5 行（表头 + 5 = A1:D6）
const ACCOUNTS = 5
check('总账：聚合结果正确', r.output.accounts === ACCOUNTS && r.output.entries_scanned === TRANSACTIONS.length, JSON.stringify({ accounts: r.output.accounts, scanned: r.output.entries_scanned }))

// 公式实际值：重算后读回活值（不是写死的数）
const rc = await call('excel_recalc', { path: P, sheet: ledgerSheet, range: 'A1:D6' })
const rows = rc.output.value
const findRow = (acc) => (Array.isArray(rows) ? rows.find((x) => x && String(x[0]) === acc) : null)
const bank = findRow('银行存款')
check('公式实际值：银行存款 借0 贷2060 余额-2060', bank && Number(bank[2]) === 2060 && Number(bank[3]) === -2060, bank ? JSON.stringify(bank) : 'row missing')
const office = findRow('办公费')
check('公式实际值：办公费 余额=1200', office && Number(office[3]) === 1200, office ? JSON.stringify(office) : 'row missing')

// ── 3. 透视表：字段与汇总正确 ─────────────────────────────────
// 借、贷两个值字段都要：银行存款只有贷方发生额，只按借方汇总会是 0
r = await call('excel_pivot_create', { path: P, sheet: '日记账', range: 'A1:E7', rows: ['科目'], values: ['借方', '贷方'] })
const ptSheet = r.output.dst_sheet
const pt = await call('excel_recalc', { path: P, sheet: ptSheet, range: 'A1:C7' })
const ptRows = Array.isArray(pt.output.value) ? pt.output.value : []
const ptFind = (acc) => ptRows.find((x) => x && String(x[0]) === acc)
const ptBank = ptFind('银行存款')
check('透视表：银行存款贷方汇总正确（2060）', ptBank && Number(ptBank[2]) === 2060, ptBank ? JSON.stringify(ptBank) : JSON.stringify(ptRows).slice(0, 200))
const ptOffice = ptFind('办公费')
check('透视表：办公费借方汇总正确（1200）', ptOffice && Number(ptOffice[1]) === 1200, ptOffice ? JSON.stringify(ptOffice) : 'row missing')
check('透视表：首次建表非复用', r.output.idempotent_reuse === false, JSON.stringify({ reuse: r.output.idempotent_reuse, name: r.output.table_name }))

r = await call('excel_pivot_refresh', { path: P })
check('透视表刷新：刷新到全部透视表', r.output.refreshed === 1, JSON.stringify(r.output))

// ── 4. 整条链重跑一遍：必须得到同一结果且不破坏工作簿 ──────────
const sheetsBefore = await raw(["excel = Officer.Excel", "excel.Workbooks.Open(a['path'])", "wb = excel.ActiveWorkbook",
  'n = [w.Name for w in wb.Worksheets]', 'wb.Close(False)', "output = json.dumps(n)"], { path: P })

r = await call('excel_journal_post', { path: P, sheet: '日记账', entries: TRANSACTIONS })
check('重跑：同批次不再重复入账', r.output.idempotent_skip === true && r.changed === false, JSON.stringify({ skip: r.output.idempotent_skip, changed: r.changed }))

r = await call('excel_ledger_gen', { path: P, journal_sheet: '日记账' })
check('重跑：总账结果一致直接跳过', r.output.idempotent_skip === true && r.changed === false, JSON.stringify({ skip: r.output.idempotent_skip }))

r = await call('excel_pivot_create', { path: P, sheet: '日记账', range: 'A1:E7', rows: ['科目'], values: ['借方', '贷方'] })
check('重跑：透视表复用同一张', r.output.idempotent_reuse === true, JSON.stringify({ reuse: r.output.idempotent_reuse }))

const sheetsAfter = await raw(["excel = Officer.Excel", "excel.Workbooks.Open(a['path'])", "wb = excel.ActiveWorkbook",
  'n = [w.Name for w in wb.Worksheets]', 'wb.Close(False)', "output = json.dumps(n)"], { path: P })
check('重跑：工作表没有堆积', JSON.stringify(sheetsBefore) === JSON.stringify(sheetsAfter), `${JSON.stringify(sheetsBefore)} -> ${JSON.stringify(sheetsAfter)}`)

const rc2 = await call('excel_recalc', { path: P, sheet: ledgerSheet, range: 'A1:D6' })
check('重跑：总账数值与首轮一致', JSON.stringify(rc2.output.value) === JSON.stringify(rc.output.value), JSON.stringify(rc2.output.value).slice(0, 120))

// ── 5. 一次异常分录测试（§八 第 6 项）+ 改正后不回潮 ────────────
const BAD = [
  { date: '2026-08-20', desc: '录入错误', account: '其他应收款', debit: 500, credit: 0 },
  { date: '2026-08-20', desc: '录入错误', account: '银行存款', debit: 0, credit: 400 },
]
r = await call('excel_journal_post', { path: P, sheet: '日记账', entries: BAD })
check('异常分录：借贷不平衡被检出', r.output.balanced === false && r.output.debit_total === 500 && r.output.credit_total === 400, JSON.stringify({ dr: r.output.debit_total, cr: r.output.credit_total }))
check('异常分录：带非致命告警而非失败', r.ok === true && r.warnings.some((w) => w.includes('借贷不平衡')), JSON.stringify(r.warnings))

const redRow = await raw([
  'excel = Officer.Excel', "excel.Workbooks.Open(a['path'])", 'wb = excel.ActiveWorkbook',
  "ws = wb.Worksheets('日记账')",
  "output = json.dumps({'color': ws.Cells(8,1).Font.Color, 'colorIndex': ws.Cells(8,1).Font.ColorIndex})",
  'wb.Close(False)',
], { path: P })
check('异常分录：错行被标红', String(redRow.color) === '255', JSON.stringify(redRow))

// 改正：replace 位置覆盖并复位颜色
const GOOD = [
  { date: '2026-08-20', desc: '录入改正', account: '其他应收款', debit: 500, credit: 0 },
  { date: '2026-08-20', desc: '录入改正', account: '银行存款', debit: 0, credit: 500 },
]
r = await call('excel_journal_post', { path: P, sheet: '日记账', entries: GOOD, on_duplicate: 'replace' })
check('改正：位置覆盖生效且已平衡', r.output.balanced === true && r.output.posted === 2, JSON.stringify(r.output))

const fixedRow = await raw([
  'excel = Officer.Excel', "excel.Workbooks.Open(a['path'])", 'wb = excel.ActiveWorkbook',
  "ws = wb.Worksheets('日记账')",
  "output = json.dumps({'color': ws.Cells(8,1).Font.Color, 'colorIndex': ws.Cells(8,1).Font.ColorIndex, 'last': ws.Cells(ws.Rows.Count,1).End(-4162).Row})",
  'wb.Close(False)',
], { path: P })
check('改正：标红被复位（不再停留 255）', String(fixedRow.color) !== '255', JSON.stringify(fixedRow))
check('改正：没有多出一行（仍是原地覆盖）', Number(fixedRow.last) === 9, `last=${fixedRow.last}`)

// 总账跟着更新，且不重复计入那一批
r = await call('excel_ledger_gen', { path: P, journal_sheet: '日记账' })
check('改正后：总账重算未跳过（内容确实变了）', r.output.idempotent_skip === false, JSON.stringify({ accounts: r.output.accounts }))
const rc3 = await call('excel_recalc', { path: P, sheet: ledgerSheet, range: 'A1:D6' })
const other = (Array.isArray(rc3.output.value) ? rc3.output.value : []).find((x) => x && String(x[0]) === '其他应收款')
check('改正后：其他应收款只计一次（余额=500）', other && Number(other[1]) === 500 && Number(other[3]) === 500, other ? JSON.stringify(other) : 'row missing')

// ── 6. 进程与文件锁：managed 链路必须把打开的工作簿都关掉 ──────
const openLeft = await raw(["excel = Officer.Excel",
  "output = json.dumps([w.Name for w in excel.Workbooks])"], {})
check('无遗留打开的工作簿（文件锁已释放）', !String(openLeft).includes('flagship-accounting.xlsx'), JSON.stringify(openLeft))

// ── 结构化验证报告（§八 第 7 项 / §五 验收指标）────────────────
const metric = (name) => REPORT.find((x) => x.metric.startsWith(name))
console.log('\n═══ 结构化验证报告 ═══')
console.log(JSON.stringify({
  task: 'office_generate_accounting_report',
  source: '脱敏交易数据（6 条分录）',
  workbook: BOOK,
  journal_total: JOURNAL_TOTAL,
  entries: TRANSACTIONS.length,
  metrics: REPORT.map(({ metric, ok }) => ({ metric, ok })),
  summary: { total: REPORT.length, passed: REPORT.length - failures, failed: failures },
}, null, 2))

console.log(failures === 0 ? '\n[FLAGSHIP] ALL PASS' : `\n[FLAGSHIP] ${failures} FAILURE(S)`)
void metric
process.exit(failures === 0 ? 0 : 1)
