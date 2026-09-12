// 注册冒烟：无 Office / officemcp 环境也能跑（CI 用）——验证 15 个工具全部注册 + 降级路径正常
// 运行: node test/register.mjs  （等价 npm test）
// 真实 COM 链路见 test/smoke.mjs / test/headless.mjs（npm run test:e2e，需本机 Office + officemcp）
import { apply, findPython, finalize, classifyError, TOOL_META, resolveCtl, parseErrMeta } from '../lib/index.mjs'

const tools = {}
const ctx = {
  tools: {
    register: (t) => {
      tools[t.name] = t
    },
  },
}
apply(ctx)

const need = [
  'office_apps', 'office_launch',
  'excel_new', 'excel_open', 'excel_read_range', 'excel_write_range',
  'excel_formula_set', 'excel_recalc', 'excel_vba_run',
  'excel_pivot_create', 'excel_pivot_refresh',
  'word_open', 'word_edit',
  'excel_journal_post', 'excel_ledger_gen',
]
const missing = need.filter((n) => !tools[n])
if (missing.length) {
  console.error(`[register] FAIL missing tools: ${missing.join(', ')}`)
  process.exit(1)
}
console.log(`[register] ${need.length} tools registered`)

// ── v0.2 统一结果契约（纯函数，不需要 Office）──
const cases = [
  // 附着模式（无 path）：改了但没落盘，saved 必须是 false
  ['attached-saved', finalize('excel_write_range', TOOL_META.excel_write_range, {}, { ok: true, output: { written: true } }),
    (r) => r.ok === true && r.changed === true && r.saved === false],
  // managed 模式（有 path）：Close 前 Save 过
  ['managed-saved', finalize('excel_write_range', TOOL_META.excel_write_range, { path: 'C:\\x.xlsx' }, { ok: true, output: {} }),
    (r) => r.saved === true],
  // 借贷不平衡：成功 + warning，但 verified 降级
  ['unbalanced', finalize('excel_journal_post', TOOL_META.excel_journal_post, { path: 'x' }, { ok: true, output: { balanced: false } }),
    (r) => r.ok === true && r.verified === false && r.warnings.length === 1],
  // 参数缺失：不可重试
  ['missing-param', finalize('excel_open', TOOL_META.excel_open, {}, { ok: false, error: '缺少 path' }),
    (r) => r.ok === false && r.error_code === 'MISSING_PARAM' && r.retryable === false],
  // 通道断：可重试
  ['channel-down', finalize('excel_read_range', TOOL_META.excel_read_range, {}, { ok: false, error: 'MCP call timeout: tools/call' }),
    (r) => r.error_code === 'CHANNEL_UNAVAILABLE' && r.retryable === true],
  // 附着模式的写失败：可能残留部分修改
  ['partial', finalize('excel_write_range', TOOL_META.excel_write_range, {}, { ok: false, error: '文件被占用' }),
    (r) => r.error_code === 'FILE_LOCKED' && r.partial_changes === true],
]
const bad = cases.filter(([, r, ok]) => !ok(r))
if (bad.length) {
  console.error(`[register] FAIL envelope: ${bad.map(([n, r]) => `${n}=${JSON.stringify(r)}`).join(' ')}`)
  process.exit(1)
}
console.log(`[register] envelope contract OK (${cases.length} cases)`)

// ── v0.3 安全动作模式契约（纯函数，不需要 Office）──
const modeCases = [
  // 四，沿用现有推断：带 path = managed，不带 = attached
  ['mode-infer-managed', resolveCtl('excel_write_range', { path: 'C:\\x.xlsx' }),
    (c) => c.mode === 'managed' && c.want_save === true && c.want_close === true],
  ['mode-infer-attached', resolveCtl('excel_write_range', {}),
    (c) => c.mode === 'attached' && c.want_save === false && c.want_close === false],
  // 显式 mode 覆盖推断
  ['mode-explicit-preview', resolveCtl('excel_write_range', { path: 'C:\\x.xlsx', mode: 'preview' }),
    (c) => c.mode === 'preview' && c.preview === true && c.want_save === false && c.want_close === true],
  // preview 永不落盘：就算显式 save:true 也不算数
  ['mode-preview-never-saves', resolveCtl('excel_ledger_gen', { mode: 'preview', save: true }),
    (c) => c.want_save === false],
  // save / close 显式优先
  ['mode-explicit-save', resolveCtl('excel_write_range', { path: 'C:\\x.xlsx', save: false }),
    (c) => c.want_save === false && c.want_close === true],
  // 非法 mode 不抛异常，交给信封出稳定码
  ['mode-invalid', resolveCtl('excel_write_range', { mode: 'overwrite' }),
    (c) => c.valid === false],
  // 只有写工具进 dry-run 集合
  ['mode-mutating-set', resolveCtl('excel_write_range', {}),
    (c) => c.mutating === true && resolveCtl('excel_read_range', {}).mutating === false],
  // preview 信封：changed/saved 恒 false
  ['preview-envelope', finalize('excel_write_range', TOOL_META.excel_write_range, { path: 'x', mode: 'preview' }, { ok: true, output: { written: true } }),
    (r) => r.ok === true && r.preview === true && r.mode === 'preview' && r.changed === false && r.saved === false],
  // 观测优先：Python 说没存，信封就不能说存了（哪怕带了 path）
  ['observed-beats-declared', finalize('excel_write_range', TOOL_META.excel_write_range, { path: 'C:\\x.xlsx' }, { ok: true, output: { written: true, saved: false } }),
    (r) => r.saved === false],
  // U0：目标文件已被用户打开 → 改了内存但没落盘，信封说实话
  ['attached-existing-open', finalize('excel_write_range', TOOL_META.excel_write_range, { path: 'C:\\x.xlsx' }, { ok: true, output: { written: true, attached_existing_open: true } }),
    (r) => r.saved === false && r.warnings.some((w) => w.includes('已被用户打开'))],
  // 非法 mode 走 MODE_INVALID，不可重试
  ['mode-invalid-envelope', finalize('excel_write_range', TOOL_META.excel_write_range, { mode: 'overwrite' }, { ok: false, error: '[MODE_INVALID] 未知 mode' }),
    (r) => r.error_code === 'MODE_INVALID' && r.retryable === false],
  // 覆盖未确认：静态码，不可重试
  ['overwrite-not-confirmed', finalize('excel_ledger_gen', TOOL_META.excel_ledger_gen, { path: 'x' }, { ok: false, error: '[OVERWRITE_NOT_CONFIRMED] 总账工作表已有内容' }),
    (r) => r.error_code === 'OVERWRITE_NOT_CONFIRMED' && r.retryable === false],
  // 保存失败：可重试（多为文件被占的瞬时原因）
  ['save-failed-retryable', finalize('excel_write_range', TOOL_META.excel_write_range, { path: 'x' }, { ok: false, error: '[SAVE_FAILED] 保存失败: 文件被占用' }),
    (r) => r.error_code === 'SAVE_FAILED' && r.retryable === true],
  // [META] 是失败元数据的唯一通道：剥掉后错误文本要干净，且 partial_changes 以观测为准
  ['err-meta-stripped', finalize('excel_write_range', TOOL_META.excel_write_range, {}, { ok: false, error: '改到一半炸了 [META]{"partial_changes": true}' }),
    (r) => r.error === '改到一半炸了' && r.partial_changes === true],
  // 高危动作未确认：静态码
  ['risky-not-confirmed', finalize('excel_vba_run', TOOL_META.excel_vba_run, { macro: 'M.X' }, { ok: false, error: '[RISKY_OP_NOT_CONFIRMED] 运行 VBA 宏是高危动作' }),
    (r) => r.error_code === 'RISKY_OP_NOT_CONFIRMED' && r.retryable === false],
  // 宏执行失败：自带码，不走模糊匹配（COM 报错文本随语言变）
  ['vba-failed', finalize('excel_vba_run', TOOL_META.excel_vba_run, { macro: 'M.X' }, { ok: false, error: '[VBA_FAILED] 宏执行失败 M.X: 无法运行宏' }),
    (r) => r.error_code === 'VBA_FAILED'],
  // 前缀码不能盖过通道错误：SSE 超时的文本里出现任何字样都不该被误分类
  ['channel-beats-prefix', finalize('excel_read_range', TOOL_META.excel_read_range, {}, { ok: false, error: 'MCP call timeout: tools/call 提到 [SAVE_FAILED] 字样' }),
    (r) => r.error_code === 'CHANNEL_UNAVAILABLE'],
]
const badMode = modeCases.filter(([, r, ok]) => !ok(r))
if (badMode.length) {
  console.error(`[register] FAIL mode: ${badMode.map(([n, r]) => `${n}=${JSON.stringify(r)}`).join(' ')}`)
  process.exit(1)
}
// parseErrMeta 的解析本身也过一遍
const pm = parseErrMeta('boom [META]{"partial_changes": true}')
if (pm.text !== 'boom' || pm.meta.partial_changes !== true) {
  console.error(`[register] FAIL parseErrMeta: ${JSON.stringify(pm)}`)
  process.exit(1)
}
console.log(`[register] mode contract OK (${modeCases.length} cases)`)

// ── 模式参数的 schema 暴露（纯函数，不需要 Office）──
// 不变量：**工具暴露 mode，就必须真的实现了 preview 分支**。暴露了却没实现，
// Agent 传 mode=preview 会拿到一次真改动——比不暴露更糟。
// 8 个会改动文档的工具全部接线完毕（preview 分支 + 幂等语义）。
const mutating = [
  'excel_formula_set', 'excel_write_range', 'excel_vba_run',
  'excel_pivot_create', 'excel_pivot_refresh',
  'excel_journal_post', 'excel_ledger_gen', 'word_edit',
]
const schemaBad = []
for (const n of mutating) {
  const p = tools[n].parameters.properties
  if (!p.mode || !p.overwrite) schemaBad.push(`${n}: 缺 ${!p.mode ? 'mode ' : ''}${!p.overwrite ? 'overwrite' : ''}`.trim())
  if (p.mode && JSON.stringify(p.mode.enum) !== JSON.stringify(['preview', 'managed', 'attached'])) schemaBad.push(`${n}: mode enum 不对`)
  if ((tools[n].parameters.required || []).includes('mode')) schemaBad.push(`${n}: mode 不该是必填（会破坏现有调用）`)
  // 只读工具不该被塞模式参数：它们没有 dry-run 差异，塞了只会让 Agent 以为能预演
}
for (const n of ['excel_read_range', 'excel_recalc', 'excel_open', 'word_open', 'office_apps']) {
  if (tools[n].parameters.properties.mode) schemaBad.push(`${n}: 只读工具不该带 mode`)
}
if (schemaBad.length) {
  console.error(`[register] FAIL schema: ${schemaBad.join(' | ')}`)
  process.exit(1)
}
console.log(`[register] schema contract OK (${mutating.length} mutating tools expose mode)`)

// 无 officemcp 时应走降级：工具可调用但返回 {ok:false} 的友好报错（而不是抛异常）
const py = findPython()
if (!py) {
  const r = await tools.office_apps.execute()
  if (r.ok) {
    console.error('[register] FAIL: 无 officemcp 时应返回 ok:false（降级），实际 ok:true')
    process.exit(1)
  }
  console.log(`[register] degraded OK: ${r.error}`)
} else {
  console.log(`[register] officemcp python 存在（${py}），跳过执行（e2e 见 npm run test:e2e）`)
}

console.log('[register] PASS')
process.exit(0) // apply() 会预热 SSE 长连接，占着事件循环不主动退出测试就永远不结束
