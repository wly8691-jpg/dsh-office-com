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

**会计旗舰**：`excel_journal_post`（写分录 + 借贷平衡校验，借≠贷标红）· `excel_ledger_gen`（日记账 → 科目总账，SUMIF 聚合 + 余额公式）

> **path 行为约定**：带 `path` 的调用，写操作（写值/公式/分录/透视/宏）执行后**保存并关闭**文件；读操作（读值/重算/打开/透视刷新）执行后直接关闭、不落盘。不带 `path` 则操作当前活动实例，不改变其开关状态。

## 统一结果协议（v0.2）

所有工具返回同一信封，Agent 只读 `ok` / `error_code` / `retryable` 就能分支，不必解析 15 种各不相同的返回形状。

```jsonc
// 成功
{ "ok": true, "operation": "excel_write_range", "changed": true, "saved": false,
  "verified": false, "warnings": [], "output": { "range": "A1:B2", "written": true } }
// 失败
{ "ok": false, "operation": "excel_open", "error_code": "MISSING_PARAM",
  "error": "缺少 path", "retryable": false, "partial_changes": false }
```

- `changed` / `saved` / `verified` 分开报：**改了 ≠ 落盘了 ≠ 核对过**。附着模式（无 `path`）改的是用户在用的工作簿，`saved` 为 `false`（还没落盘）
- `verified` 只在有校验步骤的工具上为真：`excel_recalc`（重算后读回活值）、`excel_journal_post`（借贷平衡）
- `error_code`：`CHANNEL_UNAVAILABLE`（通道断，可重试）/ `FILE_LOCKED`（文件被占，可重试）/ `MISSING_PARAM` / `DEGRADED` / `NO_ACTIVE_DOCUMENT` / `APP_UNAVAILABLE` / `EMPTY_SOURCE` / `SOURCE_RANGE_INVALID` / `OPEN_FAILED` / `UNKNOWN`
- `partial_changes`：写操作在附着模式下失败时可能已改了一半（managed 模式会在 `finally` 里 Close 丢弃，故为 `false`）

## 进程生命周期

COM 走的是本机 OfficeMCP（workbuddy Py3.13.12 自带），officemcp 的 `Officer.Excel` 是 **Dispatch 出来并永久缓存**的实例，默认不可见。

- 该 SSE 服务由插件按需 spawn（动态端口 + `~/.dsh-office-com/office.lock` singleton），宿主退出时回收
- 宿主的 python 子进程是被强杀的（`TerminateProcess`，atexit 不跑），而 Excel 只要还挂着工作簿就不会自己退 → 会留下**不可见的孤儿 EXCEL.EXE**（占内存 + 占文件锁）。插件在退出时做收尾：**启动前不存在 且 当前不可见**的实例才回收，先 `Quit()`、送不走再按 Hwnd 拿 pid 强制结束
- 用户自己开着的、或被切到可见的实例一律不动（可见 = 有人在看）。`excel_new` 会主动把实例切到可见（新建的工作簿得让用户看得见、够得着），因此它建出来的实例不参与回收
- 检查：`npm run test:leak`（需本机 Office，跑前请先关掉 Excel）。三个用例：managed 链路退出后零残留 / 可见实例不被误杀 / `excel_new` 的实例存活

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

本机无 OfficeMCP 时静默降级：插件仍正常加载（不炸 DSH 插件树），工具调用返回友好报错。

## License

MIT
