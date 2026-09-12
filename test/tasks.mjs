// tasks.mjs — 任务级工具回归（v0.4 的 6 个工具）
//   office_generate_accounting_report / office_check_workbook / office_replace_document_terms
//   office_update_monthly_report / office_apply_template / office_prepare_management_summary
//
// 与其余测试的分工：
//   register.mjs  = 纯函数契约（CI，无需 Office）
//   headless.mjs  = 底层工具编排 / 输入校验 / 返回结构
//   flagship.mjs  = 底层工具拼出的会计链路，跑两遍验可重复性
//   tasks.mjs     = **任务级工具**（一次调用跑完一整套步骤）的行为与安全闸
//   faults.mjs    = 异常路径：文件被占用 / 保存失败 / 断线重连 / 宏失败
//   leakcheck.mjs = 跨进程的 Office 进程残留
//
// 运行: node test/tasks.mjs   （需本机 Office；跑前建议关掉 Excel，免得与你自己的文档争用实例）
import { apply, runPython } from '../lib/index.mjs'
import { existsSync, rmSync, writeFileSync, statSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'

const tools = {}
apply({ tools: { register: (t) => { tools[t.name] = t } } })

let fail = 0
function check(name, cond, detail) {
  console.log(`[${cond ? 'PASS' : 'FAIL'}] ${name}${detail ? ' — ' + detail : ''}`)
  if (!cond) fail++
}

/** 走插件自己的桥做裸 COM 操作：造测试文件、独立复核结果。 */
async function raw(lines, argsObj) {
  // 路径必须归一化：Excel 对正斜杠解析异常（会把 C:/x 拼成乱路径）
  const code = ['import json, pythoncom, os', 'pythoncom.CoInitialize()',
    'a = json.loads(data) if data else {}',
    "if a.get('path'): a['path'] = a['path'].replace('/', os.sep)", ...lines].join('\n')
  const r = await runPython(code, JSON.stringify(argsObj || {}))
  if (!r.success) throw new Error(r.error)
  return r.output
}

/** 把 a 转成 Python 可用的小写 a['key'] 取值——测试代码里别写反斜杠。 */
const P = (p) => p.replace(/\\/g, '/')

const D = tmpdir()
const F = (n) => join(D, n)
const reset = (p) => { if (existsSync(p)) rmSync(p) }

// ══ A. office_generate_accounting_report ════════════════════════
{
  const CSV = F('t-tx.csv'), TPL = F('t-tpl.xlsx'), OUT = F('t-report.xlsx')
  ;[OUT].forEach(reset)
  writeFileSync(CSV, ['日期,摘要,科目,借方,贷方',
    '2026-08-03,采购办公用品,办公费,1200,0', '2026-08-03,采购办公用品,银行存款,0,1200',
    '2026-08-07,支付差旅费,差旅费,860,0', '2026-08-07,支付差旅费,银行存款,0,860',
    '2026-07-15,上月水电,水电费,300,0', '2026-07-15,上月水电,银行存款,0,300',
    ',空科目行应被跳过,,99,99'].join('\n'), 'utf8')

  await raw(['excel = Officer.Excel', 'excel.DisplayAlerts = False', 'excel.Visible = False',
    'wb = excel.Workbooks.Add()', "wb.Sheets(1).Name = '备注'",
    "wb.Sheets('备注').Range('A1').Value = '这份备注必须原样保留'",
    "wb.SaveAs(a['path'])", 'wb.Close(False)', "output = json.dumps({'ok': True})"], { path: P(TPL) })
  const tplM = statSync(TPL).mtimeMs

  let r = await tools.office_generate_accounting_report.execute({ source: CSV, output: OUT, template: TPL, period: '2026-08', mode: 'preview' })
  check('报表 preview：真预演且不落笔', r.ok && r.preview === true && !existsSync(OUT), r.ok ? `entries=${r.output?.entries}` : `[${r.error_code}] ${r.error}`)
  check('报表 preview：period 过滤 6→4、借贷合计正确', r.output?.entries === 4 && r.output?.debit_total === 2060 && r.output?.balanced === true, JSON.stringify({ e: r.output?.entries, dr: r.output?.debit_total }))
  check('报表 preview：科目数与信封口径对', r.output?.accounts === 3 && r.changed === false && r.saved === false, JSON.stringify({ n: r.output?.accounts, c: r.changed }))

  r = await tools.office_generate_accounting_report.execute({ source: CSV, output: OUT, template: TPL, period: '2026-08' })
  check('报表 managed：一次调用产出三表并落盘', r.ok && r.saved === true && existsSync(OUT) && r.output?.journal?.posted === 4 && r.output?.ledger?.accounts === 3, r.ok ? JSON.stringify(r.output?.ledger) : `[${r.error_code}] ${r.error}`)
  const v = r.output?.verification || []
  const bank = v.find((x) => x.account === '银行存款')
  check('报表：校验读回的是活公式值（银行存款 贷2060 余-2060）', bank && Number(bank.credit) === 2060 && Number(bank.balance) === -2060, bank ? JSON.stringify(bank) : 'row missing')

  check('报表：模板永不被改动（mtime 不变）', statSync(TPL).mtimeMs === tplM)
  const sheets = await raw(['excel = Officer.Excel', "excel.Workbooks.Open(a['path'])", 'wb = excel.ActiveWorkbook',
    'n = [w.Name for w in wb.Worksheets]', "note = wb.Worksheets('备注').Range('A1').Value", 'wb.Close(False)',
    "output = json.dumps({'sheets': n, 'note': note})"], { path: P(OUT) })
  check('报表：输出保留模板底稿与原内容', String(sheets.sheets).includes('备注') && String(sheets.note).includes('原样保留'), JSON.stringify(sheets))
  check('报表：输出含 日记账/总账/透视表', ['日记账', '总账', '透视表'].every((x) => String(sheets.sheets).includes(x)), String(sheets.sheets))

  r = await tools.office_generate_accounting_report.execute({ source: CSV, output: OUT, period: '2026-08' })
  check('报表：输出已存在时拒绝覆盖', !r.ok && r.error_code === 'OVERWRITE_NOT_CONFIRMED', `[${r.error_code}] ${r.error}`)
  r = await tools.office_generate_accounting_report.execute({ source: CSV, output: OUT, period: '2026-08', overwrite: true })
  check('报表：显式 overwrite 后重跑结果一致且复用同一张透视表', r.ok && r.output?.ledger?.accounts === 3 && r.output?.pivot?.table === 'PT_日记账_A1_E5', r.ok ? JSON.stringify(r.output?.pivot) : `[${r.error_code}] ${r.error}`)

  r = await tools.office_generate_accounting_report.execute({ source: CSV })
  check('报表：缺 output 报 MISSING_PARAM', !r.ok && r.error_code === 'MISSING_PARAM', `[${r.error_code}]`)
  r = await tools.office_generate_accounting_report.execute({ source: CSV, output: OUT, period: '1999-01', overwrite: true })
  check('报表：期间无记录报 EMPTY_SOURCE', !r.ok && r.error_code === 'EMPTY_SOURCE', `[${r.error_code}]`)
}

// ══ B. office_check_workbook ════════════════════════════════════
{
  const B = F('t-check.xlsx')
  reset(B)
  await raw(['excel = Officer.Excel', 'excel.DisplayAlerts = False', 'excel.Visible = False',
    'wb = excel.Workbooks.Add()', "ws = wb.Sheets(1)", "ws.Name = '数据'",
    "ws.Range('A1').Value = '科目'; ws.Range('A2').Value = '办公费'",
    "ws.Range('B1').Value = '金额'; ws.Range('B2').Value = 100",
    "ws.Range('C2').Formula = '=B2*2'",
    "ws.Range('D2').Formula = '=B2/0'",
    "ws.Range('E2').Formula = '=#REF!'",
    'ws2 = wb.Worksheets.Add()', "ws2.Name = '受保护'", "ws2.Range('A1').Value = 'x'", 'ws2.Protect()',
    "wb.SaveAs(a['path'])", 'wb.Close(False)', "output = json.dumps({'ok': True})"], { path: P(B) })
  const m0 = statSync(B).mtimeMs

  const r = await tools.office_check_workbook.execute({ path: B })
  check('体检：报出工作簿身份与两张表', r.ok && String(r.output?.path).includes('t-check') && r.output?.sheets?.length === 2, r.ok ? JSON.stringify(r.output?.sheets?.map((s) => s.name)) : `[${r.error_code}] ${r.error}`)
  check('体检：发现被保护的工作表', r.output?.sheets?.some((s) => s.name === '受保护' && s.protected === true))
  check('体检：抓到公式求值错误 ≥2 个', r.output?.formula_error_count >= 2, `count=${r.output?.formula_error_count}`)
  const cells = (r.output?.formula_errors || []).map((x) => x.cell).join(',').replace(/\$/g, '')
  check('体检：错误定位到 D2 与 E2', cells.includes('D2') && cells.includes('E2'), cells)
  check('体检：dirty 为真、safe_to_edit 为假且有可读原因', r.output?.dirty !== false && r.output?.safe_to_edit === false && (r.output?.flags || []).length >= 1, JSON.stringify(r.output?.flags))
  check('体检：只读不改文件（mtime 不变）且信封报 changed=false', statSync(B).mtimeMs === m0 && r.changed === false && r.saved === false)
  check('体检：只读工具不暴露 mode', tools.office_check_workbook.parameters.properties.mode === undefined)

  const r2 = await tools.office_check_workbook.execute({ path: B, sheet: '数据' })
  check('体检：可只检查单张表', r2.ok && r2.output?.sheets?.length === 1 && r2.output.sheets[0].name === '数据')
  const r3 = await tools.office_check_workbook.execute({})
  check('体检：缺 path 报 MISSING_PARAM', !r3.ok && r3.error_code === 'MISSING_PARAM', `[${r3.error_code}]`)
}

// ══ C. office_replace_document_terms ════════════════════════════
{
  const DOC = F('t-terms.docx'), DOC2 = F('t-terms2.docx')
  ;[DOC, DOC2].forEach(reset)
  await raw(['word = Officer.Word', 'word.DisplayAlerts = 0', 'word.Visible = False',
    'doc = word.Documents.Add()',
    "doc.Content.Text = '甲方与乙方签订合同。甲方付款，乙方交付。甲方验收，乙方开票。'",
    'rng = doc.Content', 'rng.Find.ClearFormatting()',
    "rng.Find.Execute('甲方', False, False, False, False, False, True, 0, False)",
    'rng.Font.Bold = True', "doc.SaveAs(a['path'])", 'doc.Close(False)', "output = json.dumps({'ok': True})"], { path: P(DOC) })
  const m0 = statSync(DOC).mtimeMs

  let r = await tools.office_replace_document_terms.execute({ path: DOC, mode: 'preview', terms: [{ find: '甲方', replace: '丙方' }, { find: '乙方', replace: '丁方' }] })
  const pc = r.output?.terms || []
  check('批量替换 preview：两组各命中 3 次且给上下文', r.ok && r.preview === true && pc.every((x) => x.count === 3) && typeof pc[0]?.first_hit_context === 'string', JSON.stringify(pc.map((x) => x.count)))
  check('批量替换 preview：未改文件（mtime 不变）', statSync(DOC).mtimeMs === m0 && r.changed === false)

  r = await tools.office_replace_document_terms.execute({ path: DOC, terms: [{ find: '甲方', replace: '丙方' }, { find: '乙方', replace: '丁方' }] })
  check('批量替换 managed：落盘', r.ok && r.saved === true, r.ok ? '' : `[${r.error_code}] ${r.error}`)
  const after = await raw(['word = Officer.Word', 'word.DisplayAlerts = 0', "doc = word.Documents.Open(a['path'])",
    't = doc.Content.Text', 'rng = doc.Content', 'rng.Find.ClearFormatting()',
    "rng.Find.Execute('丙方', False, False, False, False, False, True, 0, False)", 'bold = rng.Font.Bold', 'doc.Close(False)',
    "output = json.dumps({'text': t, 'bold': bold})"], { path: P(DOC) })
  const txt = String(after.text)
  check('批量替换：新词各 3 次、原词归零', (txt.match(/丙方/g) || []).length === 3 && (txt.match(/丁方/g) || []).length === 3 && !txt.includes('甲方') && !txt.includes('乙方'), txt)
  check('批量替换：**原文格式保住**（加粗还在）', after.bold === true || after.bold === -1, `Font.Bold=${JSON.stringify(after.bold)}`)

  r = await tools.office_replace_document_terms.execute({ path: DOC, terms: { 丙方: '甲方' } })
  check('批量替换：对象形式可用', r.ok && r.output?.terms?.[0]?.count === 3, JSON.stringify(r.output?.terms))

  // 计数是替换前口径：AA→BB 且 BB→CC 时，BB 不该被前一次替换影响
  await raw(['word = Officer.Word', 'word.DisplayAlerts = 0', 'doc = word.Documents.Add()',
    "doc.Content.Text = 'AA BB AA'", "doc.SaveAs(a['path'])", 'doc.Close(False)', "output = json.dumps({'ok': True})"], { path: P(DOC2) })
  r = await tools.office_replace_document_terms.execute({ path: DOC2, terms: { AA: 'BB', BB: 'CC' } })
  check('批量替换：计数是替换前口径（AA=2, BB=1）', r.output?.terms?.find((x) => x.find === 'AA')?.count === 2 && r.output?.terms?.find((x) => x.find === 'BB')?.count === 1, JSON.stringify(r.output?.terms))

  r = await tools.office_replace_document_terms.execute({ path: DOC })
  check('批量替换：缺 terms 报 MISSING_PARAM', !r.ok && r.error_code === 'MISSING_PARAM', `[${r.error_code}]`)
  r = await tools.office_replace_document_terms.execute({ path: DOC, terms: [] })
  check('批量替换：空 terms 报 MISSING_PARAM', !r.ok && r.error_code === 'MISSING_PARAM', `[${r.error_code}]`)
}

// ══ D. office_update_monthly_report ═════════════════════════════
{
  const RPT = F('t-month.xlsx'), OUT = F('t-month-out.xlsx'), CSV = F('t-aug.csv'), FOREIGN = F('t-foreign.xlsx'), FOUT = F('t-foreign-out.xlsx')
  ;[OUT, FOUT].forEach(reset)
  writeFileSync(CSV, ['日期,摘要,科目,借方,贷方', '2026-08-05,八月办公,办公费,800,0',
    '2026-08-05,八月办公,银行存款,0,800', '2026-08-09,八月差旅,差旅费,450,0',
    '2026-08-09,八月差旅,银行存款,0,450'].join('\n'), 'utf8')
  await raw(['excel = Officer.Excel', 'excel.DisplayAlerts = False', 'excel.Visible = False',
    'wb = excel.Workbooks.Add()', "ws = wb.Sheets(1)", "ws.Name = '月度数据'",
    "for _i, _h in enumerate(['期间','科目','借方','贷方']): ws.Cells(1,_i+1).Value = _h",
    "ws.Cells(2,1).Value = '2026-07'; ws.Cells(2,2).Value = '办公费'; ws.Cells(2,3).Value = 100",
    "ws.Cells(3,1).Value = '2026-07'; ws.Cells(3,2).Value = '银行存款'; ws.Cells(3,4).Value = 100",
    "wb.SaveAs(a['path'])", 'wb.Close(False)', "output = json.dumps({'ok': True})"], { path: P(RPT) })
  const m0 = statSync(RPT).mtimeMs

  let r = await tools.office_update_monthly_report.execute({ report: RPT, output: OUT, period: '2026-08', source: CSV, mode: 'preview' })
  check('月度 preview：首次该期间、其他期间将保留', r.ok && r.preview === true && r.output?.existing_rows_for_period === 0 && r.output?.other_period_rows_kept === 2 && !existsSync(OUT), r.ok ? JSON.stringify({ k: r.output?.other_period_rows_kept }) : `[${r.error_code}]`)

  r = await tools.office_update_monthly_report.execute({ report: RPT, output: OUT, period: '2026-08', source: CSV })
  check('月度 managed：写 4 保留 2、原报告 mtime 不变', r.ok && r.output?.written_rows === 4 && r.output?.kept_rows === 2 && statSync(RPT).mtimeMs === m0, r.ok ? JSON.stringify({ w: r.output?.written_rows, k: r.output?.kept_rows }) : `[${r.error_code}] ${r.error}`)
  const rows = await raw(['excel = Officer.Excel', "excel.Workbooks.Open(a['path'])", 'wb = excel.ActiveWorkbook',
    "ws = wb.Worksheets('月度数据')", 'last = ws.Cells(ws.Rows.Count,1).End(-4162).Row',
    'vals = [[ws.Cells(rr,cc).Value for cc in range(1,5)] for rr in range(1,last+1)]', 'wb.Close(False)',
    'output = json.dumps(vals)'], { path: P(OUT) })
  check('月度：输出含 7 月与 8 月共 7 行（表头+2+4）', rows.length === 7 && JSON.stringify(rows).includes('2026-07') && JSON.stringify(rows).includes('2026-08'), `rows=${rows.length}`)

  r = await tools.office_update_monthly_report.execute({ report: OUT, output: OUT, period: '2026-08', source: CSV, overwrite: true })
  check('月度：同期间重跑是**替换**不是追加（replaced=4）', r.ok && r.output?.replaced_rows === 4 && r.output?.is_first_time_period === false, r.ok ? JSON.stringify({ rep: r.output?.replaced_rows }) : `[${r.error_code}] ${r.error}`)
  const last = await raw(['excel = Officer.Excel', "excel.Workbooks.Open(a['path'])", 'wb = excel.ActiveWorkbook',
    "ws = wb.Worksheets('月度数据')", 'n = ws.Cells(ws.Rows.Count,1).End(-4162).Row', 'wb.Close(False)', 'output = json.dumps(n)'], { path: P(OUT) })
  check('月度：重跑后总行数没变（仍 7）', last === 7, `last=${last}`)

  await raw(['excel = Officer.Excel', 'excel.DisplayAlerts = False', 'wb = excel.Workbooks.Add()',
    "wb.Sheets(1).Name = '月度数据'",
    "wb.Sheets('月度数据').Range('A1').Value = '产品'; wb.Sheets('月度数据').Range('B1').Value = '销量'",
    "wb.Sheets('月度数据').Range('A2').Value = 'A'; wb.Sheets('月度数据').Range('B2').Value = 5",
    "wb.SaveAs(a['path'])", 'wb.Close(False)', "output = json.dumps({'ok': True})"], { path: P(FOREIGN) })
  r = await tools.office_update_monthly_report.execute({ report: FOREIGN, output: FOUT, period: '2026-08', source: CSV })
  check('月度：外来表格被拦住（不是月度表）', !r.ok && r.error_code === 'OVERWRITE_NOT_CONFIRMED', `[${r.error_code}] ${r.error}`)
}

// ══ E. office_apply_template ════════════════════════════════════
{
  const TPL = F('t-atpl.xlsx'), TGT = F('t-atgt.xlsx'), OUT = F('t-atpl-out.xlsx')
  reset(OUT)
  await raw(['excel = Officer.Excel', 'excel.DisplayAlerts = False', 'wb = excel.Workbooks.Add()',
    "ws = wb.Sheets(1)", "ws.Name = '月报'",
    "ws.Range('A1').Value = '项目'; ws.Range('B1').Value = '金额'",
    "ws.Range('A1:B1').Font.Bold = True; ws.Range('A1:B1').Interior.Color = 65535",
    'ws.Columns(1).ColumnWidth = 28', "ws.Range('B2').NumberFormat = '0.00%'",
    "wb.SaveAs(a['path'])", 'wb.Close(False)', "output = json.dumps({'ok': True})"], { path: P(TPL) })
  await raw(['excel = Officer.Excel', 'excel.DisplayAlerts = False', 'wb = excel.Workbooks.Add()',
    "ws = wb.Sheets(1)", "ws.Name = '月报'",
    "ws.Range('A1').Value = '项目'; ws.Range('B1').Value = '金额'",
    "ws.Range('A2').Value = '办公费'; ws.Range('B2').Value = 0.25",
    "wb.SaveAs(a['path'])", 'wb.Close(False)', "output = json.dumps({'ok': True})"], { path: P(TGT) })
  const m1 = statSync(TPL).mtimeMs, m2 = statSync(TGT).mtimeMs

  let r = await tools.office_apply_template.execute({ template: TPL, target: TGT, output: OUT, mode: 'preview' })
  check('套模板 preview：报出映射对', r.ok && r.preview === true && (r.output?.pairs || []).some((x) => x.template_sheet === '月报' && x.target_sheet === '月报'), r.ok ? JSON.stringify(r.output?.pairs) : `[${r.error_code}]`)

  r = await tools.office_apply_template.execute({ template: TPL, target: TGT, output: OUT })
  check('套模板 managed：报告 applied 项 ≥2', r.ok && (r.output?.pairs?.[0]?.applied || []).length >= 2, r.ok ? JSON.stringify(r.output?.pairs) : `[${r.error_code}] ${r.error}`)
  check('套模板：模板与目标原文件都未被改动', statSync(TPL).mtimeMs === m1 && statSync(TGT).mtimeMs === m2)

  const fmt = await raw(['excel = Officer.Excel', "excel.Workbooks.Open(a['path'])", 'wb = excel.ActiveWorkbook',
    "ws = wb.Worksheets('月报')",
    "output = json.dumps({'bold': ws.Range('A1').Font.Bold, 'fill': ws.Range('A1').Interior.Color, 'w': ws.Columns(1).ColumnWidth, 'nf': ws.Range('B2').NumberFormat, 'data': ws.Range('A2').Value, 'val': ws.Range('B2').Value})",
    'wb.Close(False)'], { path: P(OUT) })
  check('套模板：加粗/底色/列宽/数字格式全搬到位',
    (fmt.bold === true || fmt.bold === -1) && Number(fmt.fill) === 65535 && Math.round(Number(fmt.w)) === 28 && String(fmt.nf).includes('0.00%'),
    JSON.stringify(fmt))
  check('套模板：**数据一行没动**', fmt.data === '办公费' && Number(fmt.val) === 0.25, JSON.stringify({ d: fmt.data, v: fmt.val }))
}

// ══ F. office_prepare_management_summary ════════════════════════
{
  const SR = F('t-sum-src.xlsx'), OUT = F('t-sum-out.xlsx')
  reset(OUT)
  await raw(['excel = Officer.Excel', 'excel.DisplayAlerts = False', 'wb = excel.Workbooks.Add()',
    "ws = wb.Sheets(1)", "ws.Name = '明细'",
    "for _i, _h in enumerate(['期间','科目','借方']): ws.Cells(1,_i+1).Value = _h",
    "data = [('2026-07','办公费',100),('2026-07','差旅费',300),('2026-08','办公费',200),('2026-08','差旅费',500),('2026-08','水电费',300)]",
    'for _i, _row in enumerate(data):',
    '    for _j, _v in enumerate(_row): ws.Cells(2+_i, _j+1).Value = _v',
    "wb.SaveAs(a['path'])", 'wb.Close(False)', "output = json.dumps({'ok': True})"], { path: P(SR) })

  let r = await tools.office_prepare_management_summary.execute({ source: SR, output: OUT, top_n: 3, mode: 'preview' })
  check('摘要 preview：自动认列 + 合计/组数正确', r.ok && r.preview === true && r.output?.group_by === '科目' && r.output?.measure === '借方' && r.output?.total === 1400 && r.output?.groups === 3 && !existsSync(OUT), r.ok ? JSON.stringify({ g: r.output?.group_by, t: r.output?.total }) : `[${r.error_code}]`)

  r = await tools.office_prepare_management_summary.execute({ source: SR, output: OUT, top_n: 3 })
  const top = r.output?.top || []
  check('摘要 managed：排行正确（差旅费800 > 办公费300 = 水电费300）', r.ok && top[0]?.name === '差旅费' && top[0]?.value === 800, r.ok ? JSON.stringify(top) : `[${r.error_code}] ${r.error}`)
  check('摘要：占比之和为 1、识别出两个期间', Math.abs(top.reduce((s, x) => s + x.share, 0) - 1) < 1e-9 && JSON.stringify(r.output?.periods) === JSON.stringify(['2026-07', '2026-08']), JSON.stringify(r.output?.periods))

  const s = await raw(['excel = Officer.Excel', "excel.Workbooks.Open(a['path'])", 'wb = excel.ActiveWorkbook',
    "ws = wb.Worksheets('管理层摘要')",
    "output = json.dumps({'title': ws.Cells(1,1).Value, 'total': ws.Cells(5,2).Value, 'hdr': [ws.Cells(7,c).Value for c in range(1,5)], 'r8': [ws.Cells(8,c).Value for c in range(1,5)], 'nf': ws.Cells(8,4).NumberFormat})",
    'wb.Close(False)'], { path: P(OUT) })
  check('摘要表：标题/合计/表头/首行都写了', s.title === '管理层摘要' && Number(s.total) === 1400 && JSON.stringify(s.hdr) === JSON.stringify(['排名', '项目', '金额', '占比']) && s.r8[1] === '差旅费', JSON.stringify(s))
  check('摘要表：占比列有百分比格式', String(s.nf).includes('%'), String(s.nf))

  r = await tools.office_prepare_management_summary.execute({ output: OUT })
  check('摘要：缺 source 报 MISSING_PARAM', !r.ok && r.error_code === 'MISSING_PARAM', `[${r.error_code}]`)
}

// ══ G. 收尾：本测试开的文件必须都关掉了 ══════════════════════════
// 只查本测试的文件（前缀 t-）。**不能断言"整个实例没有打开的工作簿"**：
// headless / flagship 也用 excel_new 造工作簿，那些会按设计留着
// （那正是 excel_new 的语义，由 leakcheck 用例 C 守着），不该算到任务级工具头上。
// 这一条真正盯的是：任务级工具自己 Open 的文件必须自己 Close。
{
  const left = await raw(['excel = Officer.Excel',
    'output = json.dumps([w.FullName for w in excel.Workbooks])'], {})
  const mine = (Array.isArray(left) ? left : []).filter((p) => /[\\/]t-/.test(String(p)))
  check('本测试开的文件都已关闭（任务级工具自己开自己关）', mine.length === 0,
    mine.length ? JSON.stringify(mine) : `实例内其他工作簿：${JSON.stringify(left)}`)
}

console.log(fail === 0 ? '\n[TASKS] ALL PASS' : `\n[TASKS] ${fail} FAILURE(S)`)
console.log('__DONE__')
setTimeout(() => process.exit(fail === 0 ? 0 : 1), 500)
