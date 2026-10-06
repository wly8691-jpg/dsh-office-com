// kb_write_back_office 真机测试（第九批工单 W1/W2/W3 · §四 验收）：
//   T1 preview 只读打开不落笔（mtime 钉 —— 红线 2，leakcheck D 同口径）
//   T2 双闸：无 confirm 拒绝（RISKY_OP_NOT_CONFIRMED），文件不动
//   T3 写回：属性读回一致 + 内容快照逐位一致（语义等价，非字节比对）
//   T4 W2 读回闭环：kb_verify office_properties 比对（写回改指纹 → 如实报 + 重标注后 clean）
//   T5 回滚：属性删除 + 内容仍逐位一致
//   T6 kb_ 前缀校验（坏属性名不进 COM）+ 零残留
// 运行: node test/kb-writeback.mjs（需真 Excel；make-xlsx.py 经同通道生成真簿）
import { apply, findPython, cleanupOwnedApps, runPython } from '../lib/index.mjs'
import { mkdtempSync, mkdirSync, writeFileSync, statSync, rmSync } from 'node:fs'
import { spawnSync } from 'node:child_process'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { readFileSync } from 'node:fs'

const tools = {}
apply({ tools: { register: (t) => { tools[t.name] = t } } })
const wbo = tools.kb_write_back_office
const verify = tools.kb_verify
const scan = tools.kb_scan_root
const annotate = tools.kb_annotate
if (!wbo || !verify || !scan || !annotate) { console.error('[kb-wb] FAIL: 工具未注册齐'); process.exit(1) }

let failures = 0
const ok = (name, cond, detail) => {
  if (cond) console.log(`[kb-wb] OK ${name}`)
  else { failures++; console.error(`[kb-wb] FAIL ${name}${detail ? '  <- ' + detail : ''}`) }
}

const PROPS = { kb_note: '写回语义钉测试 2026-10-06', kb_sensitivity: 'private' }
const py = findPython()
if (!py) {
  console.log('[kb-wb] SKIP：本机无 officemcp python（CI 形态）')
  process.exit(0)
}

// 真簿生成（与 kb-scan 深扫回归同一通道）——红线：测试用真 xlsx，不用假容器
const sandbox = mkdtempSync(join(tmpdir(), 'kb-wb-'))
const corpus = join(sandbox, 'corpus')
mkdirSync(corpus, { recursive: true })
const book = join(corpus, 'wb-test.xlsx')
const genCode = readFileSync(new URL('./fixtures/make-xlsx.py', import.meta.url), 'utf8')
const gen = await runPython(genCode, JSON.stringify({ path: book }))
if (!gen.success) { console.error('[kb-wb] FAIL 真簿生成: ' + gen.error); process.exit(1) }
ok('真 xlsx 生成', gen.success === true)

const excelPids = () => (spawnSync('tasklist', ['/FI', 'IMAGENAME eq EXCEL.EXE', '/FO', 'CSV', '/NH'],
  { encoding: 'utf8' }).stdout || '').split('\n').map((l) => l.match(/^"[^"]+","(\d+)"/)).filter(Boolean).map((m) => Number(m[1]))
const pidBaseline = excelPids().length

// 登记（registry 指纹 = F0）
const scan0 = await scan.execute({ root: corpus })
ok('kb_scan_root 登记', scan0.ok === true && scan0.output?.stats?.files_registered === 1)
const fp0 = Object.keys(scan0.output.entries)[0]

// 标注（写回前，note/sensitivity 与 PROPS 对齐 —— W2 比对基准）
const ann = await annotate.execute({ root: corpus, fingerprint: fp0, note: PROPS.kb_note, sensitivity: 'private' })
ok('kb_annotate 登记（写回前）', ann.ok === true && ann.output?.unverifiable === false)

// ── T1 preview：只读打开，mtime 物理不变 ──
const m0 = statSync(book).mtimeMs
const pv = await wbo.execute({ path: book, properties: PROPS, mode: 'preview' })
const m1 = statSync(book).mtimeMs
ok('T1 preview 报告现值与将写值',
   pv.ok === true && pv.preview === true &&
   Array.isArray(pv.output?.properties) && pv.output.properties.length === 2 &&
   pv.output.properties.every((p) => p.exists === false) &&
   pv.output.content_sha256_before?.length === 64)
ok('T1 preview 不落笔（mtime 不变，红线 2 钉）', m1 === m0)

// ── T2 双闸：缺 confirm 不写 ──
const wNo = await wbo.execute({ path: book, properties: PROPS, mode: 'managed' })
ok('T2 无 confirm → RISKY_OP_NOT_CONFIRMED', wNo.ok === false && wNo.error_code === 'RISKY_OP_NOT_CONFIRMED')
ok('T2 拒绝后文件不动（mtime 不变）', statSync(book).mtimeMs === m0)

// ── T3 写回（managed + confirm:true）→ 关掉重开核账 ──
const w = await wbo.execute({ path: book, properties: PROPS, mode: 'managed', confirm: true })
ok('T3 写回 ok 且 applied=2（added/updated）',
   w.ok === true && Array.isArray(w.output?.applied) && w.output.applied.length === 2 &&
   w.output.applied.every((a) => a.op === 'added' || a.op === 'updated'))
ok('T3 属性读回一致（properties_after 命中）',
   w.output?.properties_after?.kb_note === PROPS.kb_note &&
   w.output?.properties_after?.kb_sensitivity === 'private')
ok('T3 内容快照逐位一致（data_identical）', w.output?.data_identical === true)
ok('T3 内容 sha 前后相等', w.output?.content_sha256_before === w.output?.content_sha256_after)
ok('T3 写回改变文件（mtime 变——§四：属性写回必然动字节，不做字节比对）', statSync(book).mtimeMs !== m0)

// ── T4 W2 读回闭环 ──
const v1 = await verify.execute({ root: corpus, office_properties: true })
ok('T4 verify 只读跑通且 office_checked=1',
   v1.ok === true && v1.output?.office_checked === 1 && v1.output?.office_verified === 1)
ok('T4 无 property-mismatch / property-missing',
   !v1.output.problems.some((p) => p.kind === 'property-mismatch' || p.kind === 'property-missing'))
ok('T4 写回改指纹 → fingerprint-changed 如实报（kb.md 语义：写回后需重新标注）',
   v1.output.problems.some((p) => p.kind === 'fingerprint-changed'))
// 重标注把新指纹钉进登记 → verify 全清
const ann2 = await annotate.execute({ root: corpus, fingerprint: fp0, note: PROPS.kb_note, sensitivity: 'private' })
ok('T4 重标注落新指纹', ann2.ok === true)
const v2 = await verify.execute({ root: corpus, office_properties: true })
ok('T4 重标注后 verify clean（含 office 属性路）',
   v2.ok === true && v2.output?.clean === true,
   JSON.stringify({ problems: v2.output?.problems, ann2: ann2.output?.annotation?.fingerprint, fp0 }))

// ── T5 回滚 ──
const rb = await wbo.execute({ path: book, properties: PROPS, rollback: true, mode: 'managed', confirm: true })
ok('T5 回滚 ok 且 applied=2（deleted）',
   rb.ok === true && rb.output?.applied?.length === 2 && rb.output.applied.every((a) => a.op === 'deleted'))
ok('T5 属性已不存在', Object.keys(rb.output?.properties_after || {}).every((k) => !k.startsWith('kb_')))
ok('T5 内容仍逐位一致（data_identical）', rb.output?.data_identical === true)

// ── T6 坏属性名不进 COM + vault 守卫 ──
const bad = await wbo.execute({ path: book, properties: { 'evil_note': 'x' }, mode: 'managed', confirm: true })
ok('T6 非 kb_ 前缀拒绝', bad.ok === false && bad.error_code === 'MISSING_PARAM')

// 零残留（DispatchEx 实例整会话 Quit）：比基线不多出 EXCEL.EXE
await new Promise((r) => setTimeout(r, 2500))
const after = excelPids().length
ok('T6 零残留（EXCEL.EXE 不多于基线）', after <= pidBaseline)
console.log(`[kb-wb] excel pids: baseline=${pidBaseline} after=${after}`)

try { await cleanupOwnedApps() } catch { /* 尽力收尾 */ }
rmSync(sandbox, { recursive: true, force: true })
if (failures) { console.error(`[kb-wb] ${failures} FAILURES`); process.exit(1) }
console.log('[kb-wb] PASS')
process.exit(0)
