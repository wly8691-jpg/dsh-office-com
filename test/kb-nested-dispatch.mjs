// kb-nested-dispatch.mjs — B2 + §1.4 的回归钉（DSH 验收 2026-10-06）。**不需要 Office**。
//
// 钉住三件事：
//   ① 名字解析：MCP 工具注册名是 `mcp__<server>__<raw>`，**直查原名必然查不到** ——
//      梯子必须还能解析到它（否则 P3-b 永远 unavailable，且那是"查错名"不是"knowlp 不在"）。
//   ② 派发形状：`ctx.tools.execute` 的入参**必须含 signal**（dsh-tools:3161 读
//      `signal.aborted`，缺了抛 `TypeError: Cannot read properties of undefined (reading 'aborted')`），
//      且含 `parent`（标嵌套派发）、`callId`（唯一）、并用**解析到的实际名字**。
//   ③ 信号透传：派发用的 signal 就是外层 exec 的 signal（不是另造一个）。
import { createHash } from 'node:crypto'
import { mkdirSync, writeFileSync, rmSync, statSync, realpathSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'

const STAMP = `kbnd-${process.pid}`
const ROOT = join(tmpdir(), `${STAMP}-root`)
const STATE = join(tmpdir(), `${STAMP}-state`)
rmSync(ROOT, { recursive: true, force: true })
rmSync(STATE, { recursive: true, force: true })
mkdirSync(ROOT, { recursive: true })
process.env.DSH_OFFICE_STATE_DIR = STATE

const { apply } = await import('../lib/index.mjs')

const rootReal = realpathSync(ROOT)
const rootId = createHash('sha256').update(rootReal).digest('hex').slice(0, 16)
const dir = join(STATE, 'kb', rootId)
mkdirSync(dir, { recursive: true })

// 一个"认不出扩展名"的文件 ⇒ 机械规则归不了类 ⇒ 必落 unassigned（语义缝才有活干）
const name = 'archive.filter'
const p = join(ROOT, name)
writeFileSync(p, 'payload')
const st = statSync(p, { bigint: true })
const fp = createHash('sha256').update(`${name}\x00${st.size}\x00${st.mtimeNs}`).digest('hex')

writeFileSync(join(dir, 'registry.json'), JSON.stringify({
  registry_version: 1, root: rootReal,
  entries: { [fp]: { source_uri: `file://${name}`, pool: 'unknown', format: 'filter',
                     size_bytes: Number(st.size), mtime_epoch: 0, fingerprint: fp,
                     reason: undefined } },
  stats: {},
}), 'utf-8')
// 分类学：label 取"knowlp 命中标题"里真实会出现的词，规则为空 ⇒ 不机械命中
writeFileSync(join(dir, 'taxonomy.json'), JSON.stringify({
  version: 1, categories: [{ id: 'cat-binghuo', label: '丙火', rules: [] }],
}), 'utf-8')

// ── 假 ctx：只认 MCP public name，**故意不认原名**（复现真实注册表）────────────
const MCP_NAME = 'mcp__knowlp__knowlp_search'
const calls = []
const signalStub = new AbortController().signal
const toolStub = { name: MCP_NAME, execute: async () => ({}) }
const tools = {}
const ctx = {
  tools: {
    register: (t) => { tools[t.name] = t },
    // 只认 MCP public name —— 复现真实注册表：**直查原名必然 undefined**
    get: (n) => (n === MCP_NAME ? toolStub : undefined),
    schemas: () => [{ name: MCP_NAME }],
    execute: async (arg) => {
      calls.push(arg)
      return { output: { hits: [{ title: '丙火女02-实况' }] } }
    },
  },
}
apply(ctx)

let fails = 0
const check = (nm, cond, detail) => {
  console.log(`  ${cond ? '[PASS]' : '[FAIL]'} ${nm}${cond ? '' : '  <- ' + detail}`)
  if (!cond) fails++
}

// 模拟 dsh 的调用形态：tool.execute(exec.arguments, exec)（dsh-tools:3310）——**两个参数**。
// reg 的包装器签名就是 (args, exec)，它再把 exec 作为第三参转给工具自己的 execute。
const exec = { callId: 'call-1', token: 'tok-1', agent: 'agent-1', signal: signalStub }
const r = await tools['kb_classify'].execute(
  { root: rootReal, use_semantics: true }, exec)

check('① 梯子解析到了 MCP public name（非直查原名）',
  r?.output?.semantics !== 'unavailable',
  'semantics=' + JSON.stringify(r?.output?.semantics) + ' seam_how=' + r?.output?.seam_how)
check('① seam_how 说明了解析途径', String(r?.output?.seam_how) === 'mcp-public',
  String(r?.output?.seam_how))
check('② 确实派发了一次', calls.length === 1, 'calls=' + calls.length)
const arg = calls[0] || {}
check('② 派发入参含 signal', arg.signal === signalStub, 'signal=' + String(arg.signal))
check('② 派发入参含 parent（标嵌套）', arg.parent === 'tok-1', 'parent=' + String(arg.parent))
check('② 派发用的是解析到的名字', arg.name === MCP_NAME, 'name=' + String(arg.name))
check('② callId 唯一且非空', typeof arg.callId === 'string' && arg.callId.length > 0, String(arg.callId))
check('② rootCallId 已带上', arg.rootCallId === 'call-1', String(arg.rootCallId))
check('① 语义真的归档了（不是空转）', (r?.output?.semantic_assignments || []).length === 1,
  JSON.stringify(r?.output?.semantic_assignments))

rmSync(ROOT, { recursive: true, force: true })
rmSync(STATE, { recursive: true, force: true })
console.log(fails ? `\nFAIL: ${fails}` : '\nPASS')
process.exit(fails ? 1 : 0)
