// 注册冒烟：无 Office / officemcp 环境也能跑（CI 用）——验证 15 个工具全部注册 + 降级路径正常
// 运行: node test/register.mjs  （等价 npm test）
// 真实 COM 链路见 test/smoke.mjs / test/headless.mjs（npm run test:e2e，需本机 Office + officemcp）
import { apply, findPython } from '../lib/index.mjs'

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
