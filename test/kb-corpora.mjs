// 四语料实测（工单 §七 验收集，只读一级扫描）——结果落 §六
import { apply } from '../lib/index.mjs'
const tools = {}
apply({ tools: { register: (t) => { tools[t.name] = t } } })
const kb = tools.kb_scan_root

const corpora = [
  ['Desktop(root)', 'C:\\Users\\wly10\\Desktop'],
  ['TradingAgents recap_data', 'D:\\TradingAgents'],
  ['Desktop 业务三项目', 'C:\\Users\\wly10\\Desktop'],
  ['a-stock-data-quant', 'D:\\a-stock-data-quant'],
]
// 语料 ③ 的三个项目目录逐个扫（与 ① 的根分开统计无意义，改为三个子目录直扫）
const batch = [
  ['①Desktop', 'C:\\Users\\wly10\\Desktop'],
  ['②TradingAgents', 'D:\\TradingAgents'],
  ['③海家电商项目', 'C:\\Users\\wly10\\Desktop\\海家电商项目'],
  ['③词元商务', 'C:\\Users\\wly10\\Desktop\\词元商务'],
  ['③混粮出口', 'C:\\Users\\wly10\\Desktop\\混粮出口'],
  ['④a-stock-data-quant', 'D:\\a-stock-data-quant'],
]
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
