// headless.mjs — 无 DSH 直接加载插件真实工具，跑四链路验收
//   excel_write_range → excel_pivot_create → excel_journal_post(balanced) → excel_ledger_gen
// 运行: node test/headless.mjs
// 环境: OFFICE_PYTHON 可选；USERPROFILE 指向可写目录（锁文件落在 <USERPROFILE>/.dsh-office-com）
import { apply } from '../lib/index.mjs'

const tools = {}
const ctx = {
  tools: {
    register: (t) => {
      tools[t.name] = t
    },
  },
}
apply(ctx)

const need = ['excel_new', 'excel_write_range', 'excel_read_range', 'excel_recalc', 'excel_pivot_create', 'excel_journal_post', 'excel_ledger_gen']
for (const n of need) {
  if (!tools[n]) throw new Error(`tool missing: ${n}`)
}

async function call(name, args) {
  const t = tools[name]
  if (!t) throw new Error(`tool not registered: ${name}`)
  const r = await t.execute(args || {})
  if (!r.ok) throw new Error(`${name} -> ${r.error}`)
  return r.output
}

let failures = 0
function check(name, cond, detail) {
  const mark = cond ? 'PASS' : 'FAIL'
  console.log(`[${mark}] ${name}${detail ? ' — ' + detail : ''}`)
  if (!cond) failures++
  return cond
}

// ── 第 1 段：写数 + 透视表（wb1）──────────────────────────
const wb1 = await call('excel_new', {})
check('excel_new', wb1.new_workbook === true && Array.isArray(wb1.sheets) && wb1.sheets.length > 0, JSON.stringify(wb1.sheets))
const s1 = wb1.sheets[0]

const source = [
  ['科目', '月份', '金额'],
  ['办公费', '1月', 100],
  ['办公费', '2月', 150],
  ['差旅费', '1月', 80],
]
const wr = await call('excel_write_range', { sheet: s1, range: 'A1:C4', value: source })
check('excel_write_range A1:C4', wr.written === true && wr.range === 'A1:C4', JSON.stringify(wr))

const rr = await call('excel_read_range', { sheet: s1, range: 'A1:C4' })
const readBack = rr.value
const sameShape = Array.isArray(readBack) && readBack.length === 4 && readBack[0].length === 3
const cellEq = (a, b) => (typeof a === 'number' && typeof b === 'number' ? Number(a) === Number(b) : String(a) === String(b))
const sameData = sameShape && readBack.every((row, i) => row.length === 3 && row.every((v, j) => cellEq(v, source[i][j])))
check('excel_read_range 回读校验', sameData, JSON.stringify(readBack))

const pv = await call('excel_pivot_create', { sheet: s1, range: 'A1:C4', rows: ['科目'], columns: ['月份'], values: ['金额'] })
check('excel_pivot_create', pv.pivot === 'ok' && typeof pv.dst_sheet === 'string', JSON.stringify(pv))
const pvSheet = pv.dst_sheet
const pvRead = await call('excel_read_range', { sheet: pvSheet, range: 'A1:F10' })
check('excel_pivot_create 透视表已渲染', Array.isArray(pvRead.value) && pvRead.value.length > 0 && JSON.stringify(pvRead.value).length > 5, JSON.stringify(pvRead.value))

// ── 第 2 段：分录 + 总账（wb2）────────────────────────────
const wb2 = await call('excel_new', {})
const s2 = wb2.sheets[0]

const jp = await call('excel_journal_post', {
  sheet: s2,
  entries: [
    { date: '2026-01-05', desc: '购买办公用品', account: '办公费', debit: 100, credit: 0 },
    { date: '2026-01-05', desc: '购买办公用品', account: '银行存款', debit: 0, credit: 100 },
  ],
})
check('excel_journal_post 借贷平衡', jp.balanced === true && jp.debit_total === 100 && jp.credit_total === 100, JSON.stringify(jp))

const lg = await call('excel_ledger_gen', { journal_sheet: s2 })
check('excel_ledger_gen 总账生成', lg.accounts === 2 && typeof lg.output_sheet === 'string', JSON.stringify(lg))
const lgSheet = lg.output_sheet

const rc = await call('excel_recalc', { sheet: lgSheet, range: 'A1:D4' })
const lgRows = Array.isArray(rc.value) ? rc.value : [rc.value]
check('excel_ledger_gen 余额公式活值', JSON.stringify(lgRows).includes('办公费') && JSON.stringify(lgRows).includes('银行存款'), JSON.stringify(lgRows))
const findRow = (rows, acc) => rows.find((r) => r && String(r[0]) === acc)
const of = findRow(lgRows, '办公费')
const bank = findRow(lgRows, '银行存款')
check('办公费 余额=100', of && Number(of[1]) === 100 && Number(of[3]) === 100, of ? JSON.stringify(of) : 'row missing')
check('银行存款 余额=-100', bank && Number(bank[2]) === 100 && Number(bank[3]) === -100, bank ? JSON.stringify(bank) : 'row missing')

// ── 附加：不平衡分录应被检出（balanced=false）────────────
const jpBad = await call('excel_journal_post', {
  sheet: s2,
  entries: [{ date: '2026-01-06', desc: '不平衡测试', account: '测试', debit: 100, credit: 90 }],
})
check('excel_journal_post 不平衡检出', jpBad.balanced === false && jpBad.debit_total === 100 && jpBad.credit_total === 90, JSON.stringify(jpBad))

console.log(failures === 0 ? '\n[HEADLESS] ALL PASS' : `\n[HEADLESS] ${failures} FAILURE(S)`)
process.exit(failures === 0 ? 0 : 1)
