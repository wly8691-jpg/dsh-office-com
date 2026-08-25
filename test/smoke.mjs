// 冒烟测试：不经过 DSH，直接验证 findPython → 起 SSE → RunPython → 真实 COM 链路
// 运行: node test/smoke.mjs
import { findPython, runPython } from '../lib/index.mjs'

async function main() {
  const py = findPython()
  console.log(`[smoke] findPython -> ${py || 'NOT FOUND'}`)
  if (!py) {
    console.error('[smoke] FAIL: officemcp python 未找到')
    process.exit(1)
  }

  // 1) SSE + RunPython 协议（无 COM 依赖）
  const r1 = await runPython("output = 'hello-from-COM'", '')
  console.log(`[smoke] 协议链路 ->`, r1)

  // 2) COM 探活（查注册/运行态，不启动 Office）
  const r2 = await runPython(
    "import json; output = json.dumps({'excel': Officer.IsAppAvailable('Excel'), 'word': Officer.IsAppAvailable('Word'), 'running': Officer.RunningApps()}, ensure_ascii=False)",
    ''
  )
  console.log(`[smoke] COM 探活 ->`, r2)

  // 3) 真实 Excel：建簿 + 写公式 + 重算取活值（四步验收第 4 步链路）
  const r3 = await runPython(
    [
      'import json, pythoncom',
      'pythoncom.CoInitialize()',
      'excel = Officer.Excel',
      "if not excel: raise Exception('Excel 不可用')",
      'book = excel.Workbooks.Add()',
      'sheet = excel.ActiveSheet',
      'sheet.Cells(1,1).Value = 2',
      'sheet.Cells(2,1).Value = 3',
      "sheet.Cells(3,1).Formula = '=SUM(A1:A2)'",
      'excel.Calculate()',
      "output = json.dumps({'sum': sheet.Cells(3,1).Value}, ensure_ascii=False)",
    ].join('\n'),
    ''
  )
  console.log(`[smoke] 真实Excel ->`, r3)

  const ok = r1.success === true && r3.success === true
  console.log(ok ? '[smoke] PASS' : '[smoke] FAIL')
  process.exit(ok ? 0 : 1)
}

main().catch((e) => {
  console.error('[smoke] EXCEPTION:', e.message)
  process.exit(1)
})
