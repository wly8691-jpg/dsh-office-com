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

**Excel 读写**：`excel_new` · `excel_open` · `excel_read_range` · `excel_write_range`

**Excel 公式 / 宏**：`excel_formula_set`（Range.Formula 原样写入）· `excel_recalc`（强制重算取活值）· `excel_vba_run`（运行已有宏）

**透视表**：`excel_pivot_create`（PivotCache/PivotTable）· `excel_pivot_refresh`

**Word**：`word_open`（打开读结构）· `word_edit`（全文查找替换）

**会计旗舰**：`excel_journal_post`（写分录 + 借贷平衡校验，借≠贷标红）· `excel_ledger_gen`（日记账 → 科目总账，SUMIF 聚合 + 余额公式）

## 安装

```bash
# npm 官方通道
dsh plugin add @bananasoldier01/dsh-office-com

# 或 GitHub 源
dsh plugin add "github:BananaSoldier01/dsh-office-com#main"
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
