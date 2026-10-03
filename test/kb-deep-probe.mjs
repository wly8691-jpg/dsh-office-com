// deep-scan focused probe (temp, not part of the suite)
import { apply, runPython, findPython, cleanupOwnedApps } from '../lib/index.mjs'
import { readFileSync, mkdirSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

const tools = {}
apply({ tools: { register: (t) => { tools[t.name] = t } } })
const kb = tools.kb_scan_root

const tree = join(tmpdir(), 'kb-deep-probe-' + Date.now())
mkdirSync(tree, { recursive: true })
const genCode = readFileSync(new URL('./fixtures/make-xlsx.py', import.meta.url), 'utf8')
const gen = await runPython(genCode, JSON.stringify({ path: join(tree, 'real.xlsx') }))
console.log('gen:', gen.success, JSON.stringify(gen.output || gen.error).slice(0, 120))

const r = await kb.execute({ root: tree, max_deep: 3 })
const d = (r.output?.deep || []).find((x) => (x.source_uri || '').endsWith('real.xlsx'))
console.log('deep entry:', JSON.stringify(d, null, 1)?.slice(0, 800))
console.log('stats:', JSON.stringify(r.output?.stats))
try { await cleanupOwnedApps() } catch {}
process.exit(0)
