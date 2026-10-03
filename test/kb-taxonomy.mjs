// kb_taxonomy 分层测试（P2-a）：增 / 改名 / 合并 / 删 + 漂移审计退出码
// 纯 JS 状态工具——不需要 Office，也不需要 COM 通道。
// 运行: node test/kb-taxonomy.mjs
import { apply } from '../lib/index.mjs'
import { mkdtempSync, mkdirSync, existsSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

const tools = {}
apply({ tools: { register: (t) => { tools[t.name] = t } } })
const get = tools.kb_taxonomy_get
const set = tools.kb_taxonomy_set
if (!get || !set) { console.error('[kb-taxonomy] FAIL: taxonomy 工具未注册'); process.exit(1) }

const sandbox = mkdtempSync(join(tmpdir(), 'kb-tax-'))
const root = join(sandbox, 'corpus')
mkdirSync(root, { recursive: true })

let failures = 0
const ok = (name, cond) => {
  if (cond) console.log(`[kb-taxonomy] OK ${name}`)
  else { failures++; console.error(`[kb-taxonomy] FAIL ${name}`) }
}

// ── add ──
let r = await set.execute({ root, ops: [
  { op: 'add', category: { id: '财务', label: '财务资料', rules: [{ kind: 'ext', pattern: '.xlsx' }] } },
  { op: 'add', category: { id: '漫剧', label: '漫剧项目', rules: [{ kind: 'name-contains', pattern: '分镜' }] } },
] })
ok('add ×2', r.ok === true && r.output.version === 1 && r.output.categories === 2)

// ── get + 审计干净 ──
r = await get.execute({ root })
ok('get 返回 taxonomy', r.ok === true && r.output.taxonomy.categories.length === 2)
ok('审计干净', r.output.audit.clean === true)

// ── 重名 add 拒绝 ──
r = await set.execute({ root, ops: [{ op: 'add', category: { id: '财务', label: '重复' } }] })
ok('重复 id 拒绝', r.ok === false && r.error_code === 'DUPLICATE_ID')

// ── rename ──
r = await set.execute({ root, ops: [{ op: 'rename', id: '漫剧', label: '漫剧与动画' }] })
r = await get.execute({ root })
const mj = r.output.taxonomy.categories.find((c) => c.id === '漫剧')
ok('rename 生效', mj?.label === '漫剧与动画')

// ── 子类别 + 删除保护 ──
r = await set.execute({ root, ops: [{ op: 'add', category: { id: '财务-发票', label: '发票', parent_id: '财务' } }] })
r = await set.execute({ root, ops: [{ op: 'delete', id: '财务' }] })
ok('有子类别删除被拒（HAS_CHILDREN）', r.ok === false && r.error_code === 'HAS_CHILDREN')

// ── merge：规则并入 + supersedes 记录 + 被并类别移除 ──
r = await set.execute({ root, ops: [
  { op: 'add', category: { id: '财务-报销', label: '报销', rules: [{ kind: 'name-contains', pattern: '报销' }] } },
] })
r = await set.execute({ root, ops: [{ op: 'merge', into_id: '财务', from_ids: ['财务-报销'] }] })
r = await get.execute({ root })
const cai = r.output.taxonomy.categories.find((c) => c.id === '财务')
const gone = !r.output.taxonomy.categories.some((c) => c.id === '财务-报销')
ok('merge：规则并入 + 类别移除',
   gone && cai.rules.some((x) => x.pattern === '报销') && cai.rules.some((x) => x.pattern === '.xlsx'))
ok('merge 记录 supersedes', (cai.supersedes || []).includes('财务-报销'))

// ── delete（无子类别，force 不需要）──
r = await set.execute({ root, ops: [{ op: 'delete', id: '财务-发票' }] })
ok('delete 叶子类别', r.ok === true)

// ── 漂移审计：重叠分类（同 kind+pattern 出现在两个类别）──
r = await set.execute({ root, ops: [
  { op: 'add', category: { id: '财务2', label: '影子财务', rules: [{ kind: 'ext', pattern: '.xlsx' }] } },
] })
r = await get.execute({ root })
const overlaps = (r.output.audit.problems || []).filter((p) => p.kind === 'overlap')
ok('漂移审计：重叠分类被点名', overlaps.length === 1 &&
   JSON.stringify(overlaps[0].categories).includes('财务') &&
   JSON.stringify(overlaps[0].categories).includes('财务2'))

// ── 悬空 parent 审计 ──
r = await set.execute({ root, ops: [
  { op: 'add', category: { id: '孤儿', label: '孤儿', parent_id: '不存在的爹' } },
] })
r = await get.execute({ root })
ok('悬空 parent 被点名',
   (r.output.audit.problems || []).some((p) => p.kind === 'dangling-parent'))

rmSync(sandbox, { recursive: true, force: true })
if (failures) { console.error(`[kb-taxonomy] ${failures} FAILURES`); process.exit(1) }
console.log('[kb-taxonomy] PASS')
process.exit(0)
