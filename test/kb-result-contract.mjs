// kb-result-contract.mjs — 两条纯判据的回归钉。**不需要 Office**。
//
// 这两条判据都属于"真机上跑了几轮才被抓到"的那类，所以抽成纯函数单测：
//   ① kbUnwrapToolResult —— 缝消费工具返回。**工具返回失败信封时 registry 不置 isError**，
//      旧实现靠 JSON.parse 渲染散文判成败，必然解析失败、被 catch 吞掉、一律当成功。
//      DSH 三轮验收 §三 抓到的就是它（"经标注写回报成功、直连同文件报失败"）。
//   ② kbFuseTrips —— B4 的"保存未落盘"保险丝。它的触发条件在正常路径不可达
//      （只有内侧核账被骗过才会走到），所以用单测钉住判据本身。

const { kbUnwrapToolResult, kbFuseTrips } = await import('../lib/index.mjs')

let fails = 0
const check = (name, cond, detail) => {
  console.log(`  ${cond ? '[PASS]' : '[FAIL]'} ${name}${cond ? '' : '  <- ' + detail}`)
  if (!cond) fails++
}
const B = (n) => BigInt(n)

// ── ① kbUnwrapToolResult ────────────────────────────────────────────────────
console.log('\n① kbUnwrapToolResult（缝消费工具返回）')
{
  // 最容易错的一条：失败信封走 createSuccessResult，isError 为 false，信封在 value 里
  const r = { isError: false, value: { ok: false, error: '属性读回不一致: kb_note' },
              content: [{ type: 'text', text: 'kb_write_back_office 失败[KB_WRITE_FAILED]: 属性读回不一致: kb_note' }] }
  const u = kbUnwrapToolResult(r)
  check('失败信封（isError=false + value.ok=false）必须判失败', u.failed === true, JSON.stringify(u))
  check('  并把工具的真实错误带出来', /属性读回不一致/.test(String(u.error)), String(u.error))
}
{
  const u = kbUnwrapToolResult({ isError: true, error: 'registry error' })
  check('isError=true 判失败', u.failed === true && /registry/.test(String(u.error)), JSON.stringify(u))
}
{
  const env = { ok: true, output: { hits: [{ title: 'x' }] } }
  const u = kbUnwrapToolResult({ isError: false, value: env, content: [{ type: 'text', text: 'ok\n{"ok":true}' }] })
  check('成功信封不判失败', u.failed === false, JSON.stringify(u))
  check('  并回传信封本体（不是 registry 外壳）', u.env === env, JSON.stringify(u.env))
}
{
  const u = kbUnwrapToolResult({ output: { ok: false, error: 'inner failed' } })
  check('退路形状 output.ok=false 也判失败', u.failed === true && /inner failed/.test(String(u.error)), JSON.stringify(u))
}
{
  const u = kbUnwrapToolResult({ value: { ok: true, output: { path: 'p' } } })
  check('纯信封（无 content）不判失败', u.failed === false && u.env.output.path === 'p', JSON.stringify(u))
}
check('空结果判失败（不静默当成功）', kbUnwrapToolResult(null).failed === true, JSON.stringify(kbUnwrapToolResult(null)))

// ── ② kbFuseTrips ───────────────────────────────────────────────────────────
console.log('\n② kbFuseTrips（B4 保险丝判据）')
const same = { size: B(100), mtimeNs: B(5000) }
const moved = { size: B(101), mtimeNs: B(6000) }

check('写 + 新增属性 + 文件一丝未动 ⇒ 触发',
  kbFuseTrips('write', { kb_note: 'a' }, {}, [], same, same).trips === true)
check('写 + 改值（写前有、值不同）+ 未动 ⇒ 触发',
  kbFuseTrips('write', { kb_note: 'new' }, { kb_note: 'old' }, [], same, same).trips === true)
check('写 + 写同值 + 未动 ⇒ 不触发（Office 会合理跳过重写）',
  kbFuseTrips('write', { kb_note: 'same' }, { kb_note: 'same' }, [], same, same).trips === false,
  JSON.stringify(kbFuseTrips('write', { kb_note: 'same' }, { kb_note: 'same' }, [], same, same)))
check('写 + 新增 + 文件确实变了 ⇒ 不触发',
  kbFuseTrips('write', { kb_note: 'a' }, {}, [], same, moved).trips === false)
check('回滚 + 确实删了属性 + 未动 ⇒ 触发',
  kbFuseTrips('rollback', { kb_note: 'a' }, { kb_note: 'a' },
    [{ name: 'kb_note', op: 'deleted' }], same, same).trips === true)
check('回滚 + 什么都没删 + 未动 ⇒ 不触发',
  kbFuseTrips('rollback', { kb_note: 'a' }, {}, [], same, same).trips === false)
check('byteChanging 如实列出触发项',
  JSON.stringify(kbFuseTrips('write', { kb_note: 'new', kb_other: 'same' },
    { kb_other: 'same' }, [], same, same).byteChanging) === JSON.stringify(['kb_note']),
  JSON.stringify(kbFuseTrips('write', { kb_note: 'new', kb_other: 'same' },
    { kb_other: 'same' }, [], same, same).byteChanging))

console.log(fails ? `\nFAIL: ${fails}` : '\nPASS')
process.exit(fails ? 1 : 0)
