# 变更记录

版本号遵循 [语义化版本](https://semver.org/lang/zh-CN/)。0.x 期间每版可能带行为变更，1.0 起以兼容性为约束。

## 1.0.1

**仅元数据与文档措辞修正，无任何代码变更。**

- 对外措辞改准：补上平台前提（Windows 本机）与"测试在哪跑"的边界。原文把"已完成回归"说得
  比实际宽——测试是在开发机上跑的，不等于在目标机器上验证过。现表述：

  > Windows 本机真实 Office 自动化插件，已完成契约、任务级、故障注入和进程泄漏回归；
  > 真实 Office 测试需要在目标机器上执行。

- 同步到 README 首段、package.json description、GitHub Release 说明与仓库 About
- 起因：npm 不允许修改已发布版本的元数据，1.0.0 的描述是冻结的，故以 patch 版修正

## 1.0.0

**Windows 本机真实 Office 自动化插件**，已完成契约、任务级、故障注入和进程泄漏回归；真实 Office 测试需要在目标机器上执行。

工单（`dsh-office-com 下一步产品工单.md`）§九 执行顺序七步全部完成。**首次稳定发布。**

### 能力

- **15 个底层工具**：应用发现/启动、Excel 新建/打开/读写、公式写入与强制重算、VBA 宏、透视表创建与刷新、Word 打开与查找替换、会计分录与科目总账
- **6 个任务级工具**：一次调用跑完一整套业务步骤，Agent 不必自己编排几十个 COM 步骤
  - `office_generate_accounting_report`：一份交易数据 → 日记账 + 借贷平衡 + 科目总账 + 透视表 + 重算 + 读回校验 + 落盘
  - `office_check_workbook`：动手前的只读体检（公式错误、外部链接、受保护表、未保存改动）
  - `office_replace_document_terms`：Word 多术语批量替换，先全部计数再统一替换，保留原文格式
  - `office_update_monthly_report`：按期间更新，同期间重跑是替换不是追加，其他期间一行不动
  - `office_apply_template`：只搬格式（字体/颜色/边框/数字格式/列宽行高/冻结窗格），数据一行不动
  - `office_prepare_management_summary`：按维度聚合 + TopN + 占比 + 期间对比

### 安全

- **三种执行模式**：`preview`（真 dry-run，真开文件看真实状态、一个字节不改）/ `managed` / `attached`
- **高风险参数必须显式**：`save` / `close` / `overwrite` / `confirm` / `refresh_external_data`，默认值一律取安全侧
- **保护用户正在编辑的工作簿**：即使给了 `path`，只要那份文件已被用户打开，就按附着处理——
  既不代关、也不代存，并在信封里报 `attached_existing_open`（否则 `Workbooks.Open` 拿到的是同一个对象，
  收尾时 `Close(False)` 会把用户的窗口关掉并丢弃其未保存修改）
- 高风险动作门控：VBA 宏需 `confirm:true`；覆盖既有内容需 `overwrite:true`

### 可靠性

- **统一结果信封**：`ok` / `operation` / `mode` / `preview` / `changed` / `saved` / `verified` /
  `warnings` / `summary` / `output`，失败另有 `error_code` / `error` / `retryable` / `partial_changes`
- **观测优先**：`changed` / `saved` / `partial_changes` 以 COM 侧的实测为准，不采信调用方的声明——
  信封不会宣称一次实际失败的保存
- **重跑安全**：同批次不重复入账、总账原地刷新、透视表复用同一张；标红每次显式设色并复位
- **进程不泄漏**：退出时回收"启动前不存在且当前不可见"的 Office 实例，用户自己开的一律不动

### 测试

| 命令 | 覆盖 | 断言数 |
|---|---|---|
| `npm test` | 工具注册 + 信封/模式/schema 契约（**无需 Office**） | 32（信封 6 + 模式 26） |
| `npm run test:e2e` | 协议链路 + 底层工具编排 | 11（headless）+ smoke |
| `npm run test:flagship` | 会计旗舰链路整条跑两遍的重复正确性 | 26 |
| `npm run test:tasks` | 6 个任务级工具的行为与安全闸 | 48 |
| `npm run test:faults` | 保存失败 / 只读 / 宏失败 / 断线重连 / 类型边界 | 19 |
| `npm run test:leak` | 跨进程的 Office 进程残留与文件锁 | 4 用例 |

> 上表数字以**实跑的 PASS 行数**为准。1.0.0 发布时这里写错过（flagship/tasks/faults
> 记成 27/45/21，实际 26/48/19；headless 记成 12，实际 11）——断言写多了写少了都是失真。

CI 只跑第一项与语法检查（其余需真实 Office）。

### 修掉的真实缺陷（都在真机上暴露，不是纸面问题）

- 只读工作簿下 `Save()` 在 `DisplayAlerts=False` 时**静默不写**，而信封报 `saved:true`——
  改为 Save 后**复核** `wb.Saved`
- 非法区域引用报 `UNKNOWN`（Excel 抛的是无文本的通用码 `0x800A03EC`）——加 `_rng()` 打稳定码
- 错误信息里漏出 pywin32 的 COM 异常元组——只取人可读描述
- `SpecialCells` 返回**多区域 Range**，相邻错格连成一片，按 Address 计数会把两个错算成一个
- 已打开的文件被当成"自己开的"关闭，会关掉用户窗口并丢弃其未保存修改

## 0.4.0

- 新增 6 个任务级工具（见上）
- 配套：`resolveCtl` 的模式推断认 `output`；`_finish` 支持"已自行 SaveAs 过就不再 Save"；
  额外打开的工作簿走 `extra_wbs` 收尾

## 0.3.0

- 安全动作模式：`preview` / `managed` / `attached`
- 控制面在 JS 侧解析一次（`resolveCtl`），Python 只消费不推断
- 旗舰链重跑安全：批次签名对账、标红复位、总账原地刷新、透视表复用
- 修掉"已打开的文件被强行关闭"（会丢掉用户未保存的修改）
- 一个版本号盖 v0.2 与 v0.3 两个里程碑（npm 上跳过从未发布的 0.2.0）

## 0.1.1

- 修 `word_edit` 格式毁损、断线重连、工作簿泄漏
- 补 LICENSE / CI / 注册测试

## 0.1.0

首个版本：COM 驱动真实 Office 实例的 DSH 原生插件，15 个底层工具。
