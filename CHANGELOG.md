# 变更记录

版本号遵循 [语义化版本](https://semver.org/lang/zh-CN/)。0.x 期间每版可能带行为变更，1.0 起以兼容性为约束。

## 1.0.3

**仅去内部标识，无功能变更。**

公开发布出去的 README 里出现过内部项目名（npm 包页上就挂着），已去除；
`lib/index.mjs` 里一条历史注释同样处理。语义一字未改，只是不点名。

- **本文件原先也把那些名字写了一遍**（为了说明"去掉了什么"）——那等于换个地方又泄露一次。
  已一并改掉。这条纪律值得记：**说明"移除了某个内部名"时，别把它写出来。**
- **git 历史不动**：按纪律历史遗留不改写；已发布版本的包也改不了，只能靠新版本覆盖页面
- 顺带做了一次全仓卫生自查（内部项目名、人名、私人路径），工作树已干净
- `dev.patch.yml` 含本机路径，但它在 `.gitignore` 里，不进仓库也不进 npm 包，无需处理

## 1.0.2

三处修复，都来自真机上的实观测——两个是真缺陷。

### 修复

- **僵尸 COM 对象不再是 `UNKNOWN`。** 背靠背跑测试时观测到 1 次的竞态：上一个进程 `Quit()` 后
  还没退干净（约 1.5s 窗口），下一个恰好 `GetActiveObject`，会绑到一个**僵尸对象**
  （进程活着、每个属性访问都抛 AttributeError）。原表现是
  `失败[UNKNOWN]: Excel.Application.Workbooks` —— 裸 pywin32 文本、不可重试、码表里查不到。
  现在：新增 `COM_OBJECT_DEAD` 码（可重试）；`runPython` 命中该形态且**服务是自己起的**时，
  连服务一起换掉（只 `resetBridge` 没用——重连还是同一个服务、`Officer.Excel` 还是缓存的那个死对象）
- **进程收尾改按 pid 判归属。** 原先 `preexisting` 记的是镜像名，于是起始时只要有一个 Excel 在跑，
  整个 `EXCEL.EXE` 就被当成"用户的"，`owned` 恒为空 → **之后自己创建的实例永不回收**。
  改为记 `{镜像名: [pid]}`，按 pid 判断"启动前不存在 → 是我们起的"
- **僵尸实例现在会被回收。** 原脚本 `if app.Visible:` 在僵尸对象上抛异常，被 `except: pass` 吞掉
  → `pid` 仍是 None → `continue`，**永不回收**；而僵尸恰恰是最该回收的那类（不可见、又占着文件锁）。
  改为两段式：先扫 COM 把可见的放进保护集、不可见的 `Quit()` 送一程；`app.Visible` 抛异常
  **不再 continue**，交给按 pid 收尾

### 新增

- `DSH_OFFICE_STATE_DIR` 环境变量：状态目录（锁文件与实例基线）的位置。**做测试隔离用这个**

### 文档

- 写明**别用改 `USERPROFILE` 的方式做隔离**：踩过，改了之后任何 `Workbook.SaveAs` 都报
  `0x800A03EC`——Excel 作为 python 子进程继承了那个假家目录，`DefaultFilePath` 变空
- 修正断言数口径漂移（flagship 27→26、tasks 45→48、faults 21→19、headless 12→11）
  与"带 mode 的工具数"（8→13）
- 进程收尾那段改为如实描述新的两层判定

### 说明

僵尸竞态**构造不出来**（试过杀掉 Excel 抽掉缓存引用、在 Quit→退出窗口里密集探测 80 次，均未复现），
所以那一项是**防御性修复**：把一次实观测到的失败模式从"无从下手"变成"有码、可重试、能自愈"，
但没有可靠的回归测试，只有分类层 3 条用例守得住。

跑 `test:leak` 前请确保**既没有残留 Excel，也没有残留 SSE 服务**——否则会读到旧格式锁而误判。

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
