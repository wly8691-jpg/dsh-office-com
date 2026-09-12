# dsh-office-com

COM 驱动**真实 Office 实例**的 DeepSeek Harness（DSH）原生插件。区别于列表内所有文件级读写的 Office 插件，本插件通过本机 OfficeMCP 的 COM 通道直接操作**运行中**的 Excel / Word：VBA 宏、透视表、公式重算、已有文档深度排版，外加一套会计旗舰场景（分录 → 总账 → 透视表 → 三表）。

## 定位差异

| 现有 Office 插件 | 本插件 |
|---|---|
| 文件级读写（openpyxl/exceljs/docx） | COM 驱动真实实例 |
| 写死数值、无法重算 | 活公式链，`Application.Calculate` 取计算后值 |
| 无 VBA / 无透视表 | 真实宏运行 + PivotCache/PivotTable 透视表 |

## 工具（15 个）

**应用发现 / 启动**：`office_apps` · `office_launch`

**Excel 读写**：`excel_new`（新建工作簿并把 Excel 切到可见，工作簿直接交到用户手上）· `excel_open` · `excel_read_range` · `excel_write_range`

**Excel 公式 / 宏**：`excel_formula_set`（Range.Formula 原样写入）· `excel_recalc`（强制重算取活值）· `excel_vba_run`（运行已有宏）

**透视表**：`excel_pivot_create`（PivotCache/PivotTable）· `excel_pivot_refresh`

**Word**：`word_open`（打开读结构）· `word_edit`（全文查找替换，真实 `Find.Execute` + `wdReplaceAll`，只改文本、保留原文格式）

**会计旗舰**：`excel_journal_post`（写分录 + 借贷平衡校验，借≠贷标红）· `excel_ledger_gen`（日记账 → 科目总账，聚合借贷 + 余额公式）

## 执行模式（v0.3）

8 个会改动文档的工具都接受 `mode`。**缺省按有无 `path` 推断**，所以老调用一字不改。

| mode | 干什么 | 什么时候用 |
|---|---|---|
| `preview` | **真经 COM 打开**文件看真实状态，回报拟变更，不落笔 | 动手前先确认「会改成什么样」 |
| `managed` | 打开 → 修改 → 校验 → 保存 → 关闭 | 有明确文件路径的批量作业 |
| `attached` | 操作当前运行实例，**不主动关闭、不自动覆盖** | 用户正开着工作簿，你想在上面干活 |

`preview` 不是参数回显，是真的打开文件读：回报真实工作簿身份、真实目标地址（如 `$A$2:$B$2`）、
会被盖掉的非空格数量、完整算好的总账、命中次数与首处上下文。跑完文件 `mtime` 一字未动
（`npm run test:leak` 用例 D 会断言这一点）。

**高风险参数必须显式**——默认值一律取安全侧：

| 参数 | 含义 | 缺省 |
|---|---|---|
| `save` | 是否落盘 | `managed` 为 `true`，`attached` 为 `false` |
| `close` | 是否关闭 | `managed` 为 `true`（只关自己开的那份）；`attached` 下无效 |
| `overwrite` | 允许覆盖既有内容 | `false`，涉及覆盖时不给就返回 `OVERWRITE_NOT_CONFIRMED` |
| `confirm` | 高危动作确认位 | `false`，`excel_vba_run` 不给就返回 `RISKY_OP_NOT_CONFIRMED` |
| `refresh_external_data` | 一并刷新外部数据连接 | `false`（它会走网络） |

**重点保护**：用户正在编辑的工作簿永不被代关、也不被代存——即使你给了 `path`，只要那份文件已经
被打开，插件会按附着处理并在信封里报 `attached_existing_open`（否则 `Workbooks.Open` 拿到的是同一个
对象，收尾时 `Close(False)` 会把用户的窗口关掉并丢弃其未保存修改）。

## 重跑语义（v0.3）

会计旗舰链路可以反复跑，第二次不会把工作簿搞坏：

- `excel_journal_post`：拿这批分录与账簿末尾逐条对账，命中同一批默认**跳过**不重复入账
  （`idempotent_skip:true`，信封如实报 `changed:false`）。真有两笔一模一样的分录用
  `on_duplicate:"append"`；改正上一次入错的分录用 `on_duplicate:"replace"`（位置覆盖并复位标红）
- `excel_ledger_gen`：总账固定落在默认表「总账」并**原地刷新**，不会每次新建一张；重算结果与表上
  一致时直接跳过。按表头判归属——是工具自己的产物就永不索要 `overwrite`，是用户数据才要
- `excel_pivot_create`：字段先对着表头校验（缺字段报 `SOURCE_RANGE_INVALID`，且不会留下孤儿工作表）；
  目标表与表名都取确定性名字，第二次跑**复用刷新**同一张（`idempotent_reuse:true`）而不是再堆一张

## 统一结果协议（v0.2 / v0.3）

所有工具返回同一信封，Agent 只读 `ok` / `error_code` / `retryable` 就能分支，不必解析 15 种各不相同的返回形状。

```jsonc
// 成功
{ "ok": true, "operation": "excel_write_range", "mode": "managed", "preview": false,
  "changed": true, "saved": true, "verified": false, "warnings": [],
  "output": { "range": "A1:B2", "written": true } }
// 预演
{ "ok": true, "operation": "excel_ledger_gen", "mode": "preview", "preview": true,
  "changed": false, "saved": false, "warnings": ["preview：未做任何修改"],
  "output": { "rows": [ { "account": "办公费", "debit": 1200, "credit": 0 } ] } }
// 失败
{ "ok": false, "operation": "excel_ledger_gen", "mode": "managed", "preview": false,
  "error_code": "OVERWRITE_NOT_CONFIRMED", "error": "…", "retryable": false, "partial_changes": false }
```

- `changed` / `saved` / `verified` 分开报：**改了 ≠ 落盘了 ≠ 核对过**。附着模式（无 `path`）改的是用户在用的工作簿，`saved` 为 `false`
- 这三个字段**以观测为准**：Python 侧知道真实发生了什么（改没改、存没存、还开着没），与声明冲突时一律信观测——信封不会宣称一次实际失败的保存
- `verified` 只在有校验步骤的工具上为真：`excel_recalc`（重算后读回活值）、`excel_journal_post`（借贷平衡）
- `error_code`：`CHANNEL_UNAVAILABLE`（通道断，可重试）/ `SAVE_FAILED`（保存失败，可重试）/ `FILE_LOCKED`（文件被占，可重试）/ `MODE_INVALID` / `OVERWRITE_NOT_CONFIRMED` / `RISKY_OP_NOT_CONFIRMED` / `VBA_FAILED` / `MISSING_PARAM` / `DEGRADED` / `NO_ACTIVE_DOCUMENT` / `APP_UNAVAILABLE` / `EMPTY_SOURCE` / `SOURCE_RANGE_INVALID` / `OPEN_FAILED` / `UNKNOWN`
- `partial_changes`：失败后工作簿是否仍停在被修改状态。只有 Python 知道（自己开的那份收尾会 Close 丢弃 → `false`；附着模式改到一半 → `true`）

## 进程生命周期

COM 走的是本机 OfficeMCP（workbuddy Py3.13.12 自带），officemcp 的 `Officer.Excel` 是 **Dispatch 出来并永久缓存**的实例，默认不可见。

- 该 SSE 服务由插件按需 spawn（动态端口 + `~/.dsh-office-com/office.lock` singleton），宿主退出时回收
- 宿主的 python 子进程是被强杀的（`TerminateProcess`，atexit 不跑），而 Excel 只要还挂着工作簿就不会自己退 → 会留下**不可见的孤儿 EXCEL.EXE**（占内存 + 占文件锁）。插件在退出时做收尾：**启动前不存在 且 当前不可见**的实例才回收，先 `Quit()`、送不走再按 Hwnd 拿 pid 强制结束
- 用户自己开着的、或被切到可见的实例一律不动（可见 = 有人在看）。`excel_new` 会主动把实例切到可见（新建的工作簿得让用户看得见、够得着），因此它建出来的实例不参与回收
- 检查：`npm run test:leak`（需本机 Office，跑前请先关掉 Excel）。四个用例：managed 链路退出后零残留 / 可见实例不被误杀 / `excel_new` 的实例存活 / `preview` 真开文件后零残留且文件 `mtime` 不变

## 跑一遍会计旗舰 Demo

一份脱敏交易数据、一次 preview、一次 managed、一次异常分录、一份结构化验证报告——
`npm run test:flagship` 就是这段（需本机 Office）：

```bash
npm run test:flagship
```

它跑的是完整业务链，且**整条跑两遍**：

```text
脱敏交易数据（6 条分录 / 5 个科目）
  ↓ preview            真开文件读真实状态，mtime 一字未动
  ↓ managed 入账        借贷平衡校验（借≠贷标红）
  ↓ 生成科目总账        余额=借-贷 活公式，重算后读回实际值
  ↓ 创建透视表          真实 PivotCache/PivotTable，按科目汇总借贷
  ↓ 刷新透视表
  ↓ 全链重跑            同批次跳过 / 总账跳过 / 透视表复用 / 工作表数与数值逐格一致
  ↓ 异常分录测试        不平衡被检出 + 标红 + 非致命告警（不是失败）
  ↓ replace 改正        位置覆盖 + 标红复位 + 总账只计一次
  ↓ 结构化验证报告      26 项指标逐条 PASS/FAIL
```

没有 OfficeMCP 时会发生什么：插件照常加载，工具调用返回带 `error_code` 的友好报错
（`DEGRADED`），不会炸 DSH 插件树。

## 测试分层

| 命令 | 层级 | 需要 Office |
|---|---|---|
| `npm test` | 工具注册 + 信封/模式/schema 契约（纯函数，CI 跑） | 否 |
| `npm run test:e2e` | 协议链路 + 工具编排（smoke / headless） | 是 |
| `npm run test:flagship` | **任务回归**：一整条业务链跑两遍的重复正确性 | 是 |
| `npm run test:leak` | 跨进程的 Office 进程残留与文件锁 | 是 |

CI 只跑第一项与语法检查——其余都要真实 Office，跑前请先关掉 Excel。

## 安装

```bash
# npm 官方通道
dsh plugin add @eqman00003/dsh-office-com

# 或 GitHub 源
dsh plugin add "github:wly8691-jpg/dsh-office-com#main"
```

## 环境要求

- Microsoft Excel / Word（真实实例）
- Python + OfficeMCP（`officemcp` 包，含 pywin32）。默认探测 workbuddy Py3.13.12，可用环境变量 `OFFICE_PYTHON` 覆盖

## 架构（三层）

1. **底层**：OfficeMCP SSE（COM 通道），`RunPython(code, data)` 万能后门驱动 `Officer.Excel/Word`
2. **中层**：自举层——探测 officemcp → 起 SSE（动态端口 + 锁文件 singleton）→ `/sse` 健康检查
3. **上层**：语义工具层 `ctx.tools.register`，把 RunPython 封装成领域工具（agent 不用写 COM 代码）

## 降级

本机无 OfficeMCP 时静默降级：插件仍正常加载（不炸 DSH 插件树）。**没有 JS 回退实现**——
所有工具都返回同一个友好报错（`error_code: DEGRADED`），只是不抛异常、不影响宿主插件树。

## License

MIT
