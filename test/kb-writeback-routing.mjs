// kb-writeback-routing.mjs — W4（工单 2026-10-06 第九批）的分支逻辑回归。
//
// 这一层**不需要 Office**：kb_annotate 的 Office 写回分支靠 host 注入的 seam
// （helpers.officeWriter），所以用假 seam 就能把路由与降级逻辑钉住。
// 真机行为（属性真写进去、内容逐格一致）由 test/kb-writeback.mjs 覆盖，那个要 Office。
//
// 钉住五件事：
//   ①  Office 文件 → 路由到 seam，且属性按 note/taxonomy_id/sensitivity 构造
//   ②  seam 报错 → 不落 write_back，只在 wb_warning 里如实说
//   ③  双闸：只有 write_back 没有 confirm → 不写
//   ④  非 .md 非 Office（.png）→ 仍走拒绝分支（新增 Office 分支没把别的格式放进来）
//   ⑤  host 没给 seam → 只登记并说明（不静默、不冒充写了）
import { createHash } from 'node:crypto'
import { mkdirSync, writeFileSync, rmSync, statSync, realpathSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'

const STAMP = `kbwb-${process.pid}`
const ROOT = join(tmpdir(), `${STAMP}-root`)
const STATE = join(tmpdir(), `${STAMP}-state`)
rmSync(ROOT, { recursive: true, force: true })
rmSync(STATE, { recursive: true, force: true })
mkdirSync(ROOT, { recursive: true })
// STATE_DIR 在模块加载时读 env，必须先设再 import
process.env.DSH_OFFICE_STATE_DIR = STATE

const { kbAnnotateImpl } = await import('../lib/kb-annotate.mjs')

const rootId = createHash('sha256').update(realpathSync(ROOT)).digest('hex').slice(0, 16)
const dir = join(STATE, 'kb', rootId)
mkdirSync(dir, { recursive: true })

// 指纹必须按**真 registry 的口径**造：registry 的指纹由 Python 侧 pool_scan 算，
// 它对 Node 的等价值是 bigint 下的 mtimeNs（B1：`st.st_mtime_ns` 在 JS 里是 undefined，
// 拿它造的"指纹"永远对不上真 registry —— 探针自己也会中这个坑）。
const mk = (name, bytes) => {
  const p = join(ROOT, name)
  writeFileSync(p, bytes)
  const st = statSync(p, { bigint: true })
  return createHash('sha256').update(`${name}\x00${st.size}\x00${st.mtimeNs}`).digest('hex')
}

const fpXlsx = mk('book.xlsx', 'fake-workbook-bytes-the-seam-is-stubbed')
const fpPng = mk('pic.png', 'png-bytes')
writeFileSync(join(dir, 'registry.json'), JSON.stringify({
  registry_version: 1, root: ROOT,
  entries: {
    [fpXlsx]: { source_uri: 'file://book.xlsx', pool: 'office', format: 'xlsx',
                size_bytes: 1, mtime_epoch: 1, fingerprint: fpXlsx },
    [fpPng]: { source_uri: 'file://pic.png', pool: 'image', format: 'png',
               size_bytes: 1, mtime_epoch: 1, fingerprint: fpPng },
  },
  stats: {},
}), 'utf-8')

let fails = 0
const check = (name, cond, detail) => {
  console.log(`  ${cond ? '[PASS]' : '[FAIL]'} ${name}${cond ? '' : '  <- ' + detail}`)
  if (!cond) fails++
}

let seen = null
const okWriter = async (abs, props) => { seen = { abs, props }; return { ok: true } }

let r = await kbAnnotateImpl(
  { root: ROOT, fingerprint: fpXlsx, note: '测试标注', taxonomy_id: 'cat-1',
    sensitivity: 'commercial', write_back: true, confirm: true },
  { officeWriter: okWriter })
check('① Office 路由到 seam', seen !== null, 'seam 未被调用')
check('① 属性构造正确',
  seen && seen.props.kb_note === '测试标注' && seen.props.kb_taxonomy === 'cat-1' &&
  seen.props.kb_sensitivity === 'commercial', JSON.stringify(seen && seen.props))
check('① 落库 write_back=true', r.output.write_back === true, JSON.stringify(r.output))
check('① 无 wb_warning', !r.output.wb_warning, String(r.output.wb_warning))

const badWriter = async () => ({ error: 'seam-unavailable：拿不到审批通道' })
r = await kbAnnotateImpl(
  { root: ROOT, fingerprint: fpXlsx, note: 'x', write_back: true, confirm: true },
  { officeWriter: badWriter })
check('② seam 失败 -> 不落 write_back', r.output.write_back === false, JSON.stringify(r.output))
check('② seam 失败 -> 有 wb_warning', /Office 写回失败/.test(String(r.output.wb_warning)),
  String(r.output.wb_warning))

r = await kbAnnotateImpl(
  { root: ROOT, fingerprint: fpXlsx, note: 'x', write_back: true },
  { officeWriter: okWriter })
check('③ 缺 confirm -> 不写回', r.output.write_back === false, JSON.stringify(r.output))
check('③ 缺 confirm -> 有说明（B6：被闸挡住要说出来）',
  /双闸需 write_back:true 且 confirm:true/.test(String(r.output.wb_warning)),
  String(r.output.wb_warning))

seen = null
r = await kbAnnotateImpl(
  { root: ROOT, fingerprint: fpPng, note: 'x', write_back: true, confirm: true },
  { officeWriter: okWriter })
check('④ .png 仍走拒绝分支',
  seen === null && /仅支持 .md 前置元数据与 Office/.test(String(r.output.wb_warning)),
  String(r.output.wb_warning))

r = await kbAnnotateImpl(
  { root: ROOT, fingerprint: fpXlsx, note: 'x', write_back: true, confirm: true }, {})
check('⑤ 无 seam -> 只登记并说明',
  r.output.write_back === false && /需 host 提供审批 seam/.test(String(r.output.wb_warning)),
  String(r.output.wb_warning))

// ⑥ B1 回归钉（DSH 验收 2026-10-06）：Node 侧算的指纹必须与 registry（Python 算）**相等**。
// 原实现取 `st.st_mtime_ns`（Python 的属性名）在 JS 里恒 undefined ⇒ 指纹永远对不上 ⇒
// 标注恒 unverifiable、kb_verify 恒报 fingerprint-changed、写回后重算的"新指纹"同样是错的。
r = await kbAnnotateImpl({ root: ROOT, fingerprint: fpXlsx, note: 'B1 pin' }, {})
check('⑥ B1：完好文件 unverifiable=false', r.output.annotation.unverifiable === false,
  'unverifiable=' + r.output.annotation.unverifiable)
check('⑥ B1：落库指纹 == registry 指纹', r.output.annotation.fingerprint === fpXlsx,
  r.output.annotation.fingerprint + ' vs ' + fpXlsx)

rmSync(ROOT, { recursive: true, force: true })
rmSync(STATE, { recursive: true, force: true })
console.log(fails ? `\nFAIL: ${fails}` : '\nPASS')
process.exit(fails ? 1 : 0)
