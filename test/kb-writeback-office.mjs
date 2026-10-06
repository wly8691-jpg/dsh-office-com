// kb_write_back_office 真机测试（第十批 · B4 根因单）：
//   B4 的核心：Word / PowerPoint 的自定义属性此前【不落盘】——
//   Add 成功、Save() 不抛异常、包却没变。根因见 lib/index.mjs 的 _persist 注释。
//
//   T1 docx 落盘：写后 docProps/custom.xml 出现 + 文件字节变了 + 内容快照一致
//   T2 pptx 落盘：同上
//   T3 独立只读复核：另起一次 RunPython 打开包，确认属性真的写在磁盘上
//   T4 回滚：属性消失 + 内容仍逐位一致
//   T5 xlsx 回归：Excel 路径不许被这次改动弄坏
//   T6 保险丝 kbFuseTrips 判据（含"改值"——DSH 三轮验收指出的盲区）
//   T7 零残留：收尾后 WINWORD / POWERPNT / EXCEL 不多于基线
//
// 运行: node test/kb-writeback-office.mjs（需真 Word + PowerPoint）
import { apply, findPython, cleanupOwnedApps, runPython, kbFuseTrips } from '../lib/index.mjs'
import { mkdtempSync, mkdirSync, statSync, rmSync } from 'node:fs'
import { spawnSync } from 'node:child_process'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { readFileSync } from 'node:fs'

const tools = {}
apply({ tools: { register: (t) => { tools[t.name] = t } } })
const wbo = tools.kb_write_back_office
if (!wbo) { console.error('[kb-wb-office] FAIL: kb_write_back_office 未注册'); process.exit(1) }

let failures = 0
const ok = (name, cond, extra) => {
  if (cond) console.log(`[kb-wb-office] OK ${name}`)
  else { failures++; console.error(`[kb-wb-office] FAIL ${name}${extra ? ' :: ' + extra : ''}`) }
}

const py = findPython()
if (!py) { console.log('[kb-wb-office] SKIP：本机无 officemcp python（CI 形态）'); process.exit(0) }

const NOTE = 'B4 落盘钉测试 2026-10-06'
const PROPS = { kb_note: NOTE, kb_sensitivity: 'private' }

// ── fixture：现造真文件（不依赖任何私人路径）──
const MAKE = `
import json, os
args = json.loads(data) if data else {}
import pythoncom
from win32com.client import DispatchEx
# COM apartment must be initialised on this thread before any DispatchEx. The
# plugin's own blobs do it in their preamble (index.mjs:811 spells out the
# -2147221008 you get otherwise); this snippet is a bare RunPython payload, so
# it has to do it itself. Without it the test only passes when some earlier
# blob already initialised the apartment in the same SSE process -- i.e. it is
# order-dependent and fails on a cold process. Reproduced 2026-10-07: T1/T2
# failed with CO_E_NOTINITIALIZED while T5 passed in the same run.
pythoncom.CoInitialize()
p = args['path']; ext = os.path.splitext(p)[1].lower()
if ext == '.docx':
    a = DispatchEx('Word.Application'); a.Visible = False; a.DisplayAlerts = False
    d = a.Documents.Add(); d.Content.Text = 'kb b4 fixture'
    d.SaveAs2(p, 16); d.Close(False); a.Quit()
elif ext == '.pptx':
    a = DispatchEx('PowerPoint.Application')
    d = a.Presentations.Add(); d.Slides.Add(1, 12)
    d.SaveAs(p, 24); d.Close(); a.Quit()
else:
    raise Exception('unsupported fixture ext: %s' % ext)
output = json.dumps({'ok': True})
`

// ── 独立只读复核：打开包看 docProps/custom.xml 到底在不在、值对不对 ──
const INSPECT = `
import json, os, zipfile, re
args = json.loads(data) if data else {}
p = args['path']
st = os.stat(p)
custom = None; values = {}
try:
    with zipfile.ZipFile(p) as z:
        if 'docProps/custom.xml' in z.namelist():
            custom = True
            x = z.read('docProps/custom.xml').decode('utf-8', 'replace')
            for m in re.finditer(r'<property[^>]*name="([^"]*)"[^>]*>(.*?)</property>', x, re.S):
                t = re.search(r'<vt:lpwstr>(.*?)</vt:lpwstr>', m.group(2), re.S)
                values[m.group(1)] = t.group(1) if t else m.group(2)
        else:
            custom = False
except Exception as e:
    custom = 'ERR:%s' % e
output = json.dumps({'size': st.st_size, 'mtime_ns': st.st_mtime_ns, 'custom': custom, 'values': values})
`

const inspect = async (path) => {
  const r = await runPython(INSPECT, JSON.stringify({ path }))
  if (!r.success) throw new Error('inspect failed: ' + r.error)
  return r.output
}

const pids = (image) => (spawnSync('tasklist', ['/FI', `IMAGENAME eq ${image}`, '/FO', 'CSV', '/NH'],
  { encoding: 'utf8' }).stdout || '').split('\n').map((l) => l.match(/^"[^"]+","(\d+)"/)).filter(Boolean).map((m) => Number(m[1]))
const officePids = () => pids('WINWORD.EXE').length + pids('POWERPNT.EXE').length + pids('EXCEL.EXE').length

const sandbox = mkdtempSync(join(tmpdir(), 'kb-wb-office-'))
const pidBaseline = officePids()

// ── T1/T2: docx + pptx 真落盘 ──
for (const [label, name] of [['T1 docx', 'wb-office.docx'], ['T2 pptx', 'wb-office.pptx']]) {
  const f = join(sandbox, name)
  const mk = await runPython(MAKE, JSON.stringify({ path: f }))
  ok(`${label} fixture 造出真文件`, mk.success === true, mk.error)
  if (!mk.success) continue

  const before = await inspect(f)
  ok(`${label} 起点无 custom.xml`, before.custom === false, JSON.stringify(before.custom))

  const w = await wbo.execute({ path: f, properties: PROPS, mode: 'managed', confirm: true })
  ok(`${label} 写回 ok`, w.ok === true, w.error)
  ok(`${label} 属性读回一致（properties_after 命中）`,
     w.output?.properties_after?.kb_note === NOTE, JSON.stringify(w.output?.properties_after))
  ok(`${label} 内容快照逐位一致（data_identical）`, w.output?.data_identical === true)
  ok(`${label} 内容 sha 前后相等`, w.output?.content_sha256_before === w.output?.content_sha256_after)

  // ★ B4 的要害：包必须真的变了
  const after = await inspect(f)
  ok(`${label} ★ docProps/custom.xml 出现（B4 的病根）`, after.custom === true, JSON.stringify(after.custom))
  ok(`${label} ★ 文件字节变了`, after.size !== before.size || after.mtime_ns !== before.mtime_ns,
     `before=${before.size}/${before.mtime_ns} after=${after.size}/${after.mtime_ns}`)
  ok(`${label} ★ 独立只读复核读到属性值`, after.values?.kb_note === NOTE, JSON.stringify(after.values))

  // ── T3 回滚 ──
  const rb = await wbo.execute({ path: f, properties: PROPS, rollback: true, mode: 'managed', confirm: true })
  ok(`${label} T3 回滚 ok`, rb.ok === true, rb.error)
  ok(`${label} T3 内容仍逐位一致`, rb.output?.data_identical === true)
  const rolled = await inspect(f)
  ok(`${label} T3 独立复核属性已消失`,
     !rolled.values || rolled.values.kb_note === undefined, JSON.stringify(rolled.values))
}

// ── T5 xlsx 回归（Excel 路径不许被这次改动弄坏）──
{
  const book = join(sandbox, 'wb-office.xlsx')
  const genCode = readFileSync(new URL('./fixtures/make-xlsx.py', import.meta.url), 'utf8')
  const gen = await runPython(genCode, JSON.stringify({ path: book }))
  ok('T5 xlsx fixture 生成', gen.success === true, gen.error)
  if (gen.success) {
    const before = await inspect(book)
    const w = await wbo.execute({ path: book, properties: PROPS, mode: 'managed', confirm: true })
    ok('T5 xlsx 写回 ok', w.ok === true, w.error)
    ok('T5 xlsx data_identical', w.output?.data_identical === true)
    const after = await inspect(book)
    ok('T5 xlsx custom.xml 出现', after.custom === true)
    ok('T5 xlsx 独立复核读到属性值', after.values?.kb_note === NOTE, JSON.stringify(after.values))
    const rb = await wbo.execute({ path: book, properties: PROPS, rollback: true, mode: 'managed', confirm: true })
    ok('T5 xlsx 回滚 ok 且内容一致', rb.ok === true && rb.output?.data_identical === true, rb.error)
  }
}

// ── T6 保险丝判据（含"改值"盲区）──
{
  const same = { size: 100, mtimeNs: 1n }
  const bigger = { size: 200, mtimeNs: 2n }
  ok('T6 新增属性 + 文件未变 → 保险丝跳',
     kbFuseTrips('write', { kb_note: 'a' }, {}, [], same, same).trips === true)
  ok('T6 改值（不同值）+ 文件未变 → 保险丝跳（DSH 三轮指出的盲区已补）',
     kbFuseTrips('write', { kb_note: 'new' }, { kb_note: 'old' }, [], same, same).trips === true)
  ok('T6 写同值 + 文件未变 → 不跳（Office 会合理跳过重写，不能误报）',
     kbFuseTrips('write', { kb_note: 'same' }, { kb_note: 'same' }, [], same, same).trips === false)
  ok('T6 新增属性 + 文件确实变了 → 不跳',
     kbFuseTrips('write', { kb_note: 'a' }, {}, [], same, bigger).trips === false)
  ok('T6 回滚删属性 + 文件未变 → 保险丝跳',
     kbFuseTrips('rollback', {}, {}, [{ name: 'kb_note', op: 'deleted' }], same, same).trips === true)
  ok('T6 无必然改字节的改动 + 文件未变 → 不跳',
     kbFuseTrips('write', {}, {}, [], same, same).trips === false)
}

// ── T7 零残留 ──
await new Promise((r) => setTimeout(r, 2500))
const pidAfter = officePids()
ok('T7 零残留（Office 进程不多于基线）', pidAfter <= pidBaseline, `baseline=${pidBaseline} after=${pidAfter}`)
console.log(`[kb-wb-office] office pids: baseline=${pidBaseline} after=${pidAfter}`)

try { await cleanupOwnedApps() } catch { /* 尽力收尾 */ }
rmSync(sandbox, { recursive: true, force: true })
if (failures) { console.error(`[kb-wb-office] ${failures} FAILURES`); process.exit(1) }
console.log('[kb-wb-office] PASS')
process.exit(0)
