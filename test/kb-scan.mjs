// kb_scan_root 分层测试（工单 2026-10-03 §五-2）：
//   守卫（纯 JS，降级态也必须生效）：vault 重叠 → VAULT_FORBIDDEN，绝不进 COM
//   合成树（需 COM 通道）：分类/指纹稳定/不分类两档/artifact 只落状态目录
//   VBA 通道：本机信任中心未开 → 如实 VBA_ACCESS_DENIED（这是它今天的真实行为，不是错误）
// 运行: node test/kb-scan.mjs
import { apply, findPython, cleanupOwnedApps } from '../lib/index.mjs'
import { mkdtempSync, mkdirSync, writeFileSync, existsSync, readdirSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

const tools = {}
apply({ tools: { register: (t) => { tools[t.name] = t } } })
const kb = tools.kb_scan_root
if (!kb) { console.error('[kb-scan] FAIL: kb_scan_root 未注册'); process.exit(1) }

let failures = 0
const ok = (name, cond) => {
  if (cond) console.log(`[kb-scan] OK ${name}`)
  else { failures++; console.error(`[kb-scan] FAIL ${name}`) }
}

// ── 守卫（红线 2，不需要 COM）────────────────────────────────
const sandbox = mkdtempSync(join(tmpdir(), 'kb-scan-'))
const fakeVault = join(sandbox, 'vault')
const fakeGraph = join(fakeVault, '系统', 'knowlp-graph')
mkdirSync(fakeGraph, { recursive: true })
writeFileSync(join(fakeGraph, 'dual_graph.json'), '{}', 'utf-8')

process.env.KNOWLP_VAULT = fakeVault
{
  const r = await kb.execute({ root: fakeVault })            // root == vault
  ok('guard: root==vault 拒绝', r.ok === false && r.error_code === 'VAULT_FORBIDDEN')
}
{
  const r = await kb.execute({ root: fakeGraph })            // root 在 vault 内
  ok('guard: root 在 vault 内拒绝', r.ok === false && r.error_code === 'VAULT_FORBIDDEN')
}
{
  const outer = mkdtempSync(join(tmpdir(), 'kb-outer-'))     // root 包含 vault
  const r = await kb.execute({ root: sandbox })              // sandbox 含 fakeVault
  ok('guard: root 包含 vault 拒绝', r.ok === false && r.error_code === 'VAULT_FORBIDDEN')
  rmSync(outer, { recursive: true, force: true })
}
{
  const r = await kb.execute({ root: join(sandbox, '不存在目录') })
  ok('root 不存在 → OPEN_FAILED（不进 COM）', r.ok === false && r.error_code === 'OPEN_FAILED')
}

// ── 合成树（需 COM 通道；无通道则如实跳过，不假装跑过）──────────
const tree = join(sandbox, 'corpus')
mkdirSync(join(tree, 'sub'), { recursive: true })
writeFileSync(join(tree, 'a.md'), 'hello', 'utf-8')
writeFileSync(join(tree, 'sub', 'b.md'), 'world', 'utf-8')
writeFileSync(join(tree, 'conf.sample'), 'cfg', 'utf-8')
writeFileSync(join(tree, 'junk.pyc'), 'compiled', 'utf-8')
writeFileSync(join(tree, 'weird.xyz'), 'mystery', 'utf-8')
writeFileSync(join(tree, 'noext-text'), 'plain', 'utf-8')
writeFileSync(join(tree, 'noext-bin'), '\x00\x01\x02', 'utf-8')
writeFileSync(join(tree, '.env'), 'secret', 'utf-8')          // 点文件：整类跳过
writeFileSync(join(tree, 'sheet.xlsx'), 'PK\x03\x04 fake', 'utf-8')

const py = findPython()
if (!py) {
  console.log('[kb-scan] SKIP 合成树段：本机无 officemcp python（CI 形态），守卫段已覆盖')
} else {
  const r1 = await kb.execute({ root: tree })
  if (!r1.ok) {
    console.log(`[kb-scan] SKIP 合成树段：通道不可用（${r1.error_code}: ${r1.error?.slice(0, 80)}）`)
  } else {
    const out = r1.output
    const list = Object.values(out.entries || {})
    const byUri = (u) => list.find((e) => e.source_uri === u)
    ok('md → text', byUri('file://a.md')?.pool === 'text')
    ok('嵌套目录相对路径正确', byUri('file://sub/b.md')?.pool === 'text')
    ok('.sample → code（P1.5-6）', byUri('file://conf.sample')?.pool === 'code')
    ok('.pyc → not-material 档（P1.5-2）',
       byUri('file://junk.pyc')?.pool === 'not-material' &&
       byUri('file://junk.pyc')?.tier === 'not-material')
    ok('.xyz → unrecognized + 原因码 unclaimed-extension',
       byUri('file://weird.xyz')?.pool === 'unknown' &&
       byUri('file://weird.xyz')?.reason === 'unclaimed-extension')
    ok('无扩展名文本 → text（magic）', byUri('file://noext-text')?.pool === 'text')
    ok('无扩展名二进制 → unknown + no-extension',
       byUri('file://noext-bin')?.pool === 'unknown' &&
       byUri('file://noext-bin')?.reason === 'no-extension')
    ok('点文件整类跳过（.env 不登记）',
       !list.some((e) => e.source_uri.endsWith('.env')))

    // 指纹稳定：跑两遍，身份集必须全等（验收：重复扫描不产生多个池身份）
    const r2 = await kb.execute({ root: tree })
    const ids1 = Object.values(out.entries).map((e) => e.fingerprint).sort()
    const ids2 = Object.values(r2.output?.entries || {}).map((e) => e.fingerprint).sort()
    ok('重复扫描身份集全等', JSON.stringify(ids1) === JSON.stringify(ids2))

    // 不分类两档分列 + 排序（unrecognized 在前）
    const u = out.unclassified || {}
    ok('unrecognized 头条非空且含 .xyz', (u.unrecognized || []).some((e) => e.source_uri.endsWith('weird.xyz')))
    ok('not-material 尾段含 .pyc', (u.not_material || []).some((e) => e.source_uri.endsWith('junk.pyc')))

    // artifact 只落状态目录，绝不写被测目录
    const art = out.artifacts
    ok('registry 落状态目录', art?.registry && existsSync(art.registry) && art.registry.includes(join('.dsh-office-com', 'kb')))
    ok('unclassified 落状态目录', art?.unclassified && existsSync(art.unclassified))
    const treeFiles = readdirSync(tree)
    ok('被测目录零写入（无 JSON 残留）', !treeFiles.some((f) => f.endsWith('.json')))

    // VBA 通道：如实上报（本机信任中心未开 → VBA_ACCESS_DENIED 是预期值）
    console.log(`[kb-scan] vba_channel = ${out.vba_channel}（本机实测值；未开 AccessVBOM 时 VBA_ACCESS_DENIED 即正确行为）`)
    ok('vba_channel 字段存在', typeof out.vba_channel === 'string')

    // 深扫：max_deep>0 会真开 Excel——本机 xlsx 是假容器，预期逐簿软失败不崩
    const rd = await kb.execute({ root: tree, max_deep: 2 })
    ok('深扫路径不崩（假 xlsx 软失败）', rd.ok === true && Array.isArray(rd.output?.deep))
  }
}

rmSync(sandbox, { recursive: true, force: true })
// SSE 服务子进程会挂着事件循环——断言结束后显式收尾并退出
try { await cleanupOwnedApps() } catch { /* 尽力收尾 */ }
if (failures) { console.error(`[kb-scan] ${failures} FAILURES`); process.exit(1) }
console.log('[kb-scan] PASS')
process.exit(0)
