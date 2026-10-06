// 四语料实测（工单 §七 验收集，只读一级扫描）——结果落 §六
//
// ⚠️ 语料清单**不进仓**：真实路径与本机业务目录名都是私人数据（公开仓红线）。
// 照本仓已有的 `.local.*` 先例：清单放 `test/corpora.local.json`（已 gitignore），
// 格式 `[["标签", "绝对路径"], ...]`。文件不在就跳过 —— 跳过不是失败。
import { existsSync, readFileSync } from 'node:fs'
import { apply } from '../lib/index.mjs'
const tools = {}
apply({ tools: { register: (t) => { tools[t.name] = t } } })
const kb = tools.kb_scan_root

const LIST_PATH = new URL('./corpora.local.json', import.meta.url)
if (!existsSync(LIST_PATH)) {
  console.log('[kb-corpora] 跳过：没有 test/corpora.local.json（私人语料清单，不入仓）')
  console.log('  格式：[["标签", "绝对路径"], ...]')
  process.exit(0)
}
const batch = JSON.parse(readFileSync(LIST_PATH, 'utf8'))
for (const [name, root] of batch) {
  const t0 = Date.now()
  const r = await kb.execute({ root, max_files: 50000 })
  const ms = Date.now() - t0
  if (!r.ok) { console.log(`${name}: FAIL ${r.error_code} ${String(r.error).slice(0, 60)}`); continue }
  const o = r.output
  const pools = {}
  for (const e of Object.values(o.entries)) pools[e.pool] = (pools[e.pool] || 0) + 1
  console.log(`${name}: ${ms}ms files=${o.stats.files_registered} pools=${JSON.stringify(pools)} unrec=${o.unclassified.unrecognized.length} notmat=${o.unclassified.not_material.length} art=${o.artifacts.dir}`)
}
process.exit(0)
