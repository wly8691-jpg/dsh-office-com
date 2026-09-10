// 注册冒烟：无 Office / officemcp 环境也能跑（CI 用）——验证 15 个工具全部注册 + 降级路径正常
// 运行: node test/register.mjs  （等价 npm test）
// 真实 COM 链路见 test/smoke.mjs / test/headless.mjs（npm run test:e2e，需本机 Office + officemcp）
import { apply, findPython, finalize, classifyError, TOOL_META } from '../lib/index.mjs'

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
