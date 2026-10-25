# 变更记录

版本号遵循 [语义化版本](https://semver.org/lang/zh-CN/)。0.x 期间每版可能带行为变更，1.0 起以兼容性为约束。

## 1.2.2

**诊断向修复：把「插件坏了」和「workbuddy 升级把包装丢了」分开**（实机事故：workbuddy 升级
Python 后整包换掉 site-packages，officemcp / pywin32 消失，插件按设计降级，但报错只有一句
"need OfficeMCP installed"，把排查引向重装插件）。

- `DEGRADED` 的报错改为 `findPythonDiagnostics()`：逐个列出探过的解释器与各自缺失的模块
  （officemcp / pythoncom / win32com.client …），并给出两条可直接复制的补救路径。
- 升级备份 `*.old.<pid>` 只作为**线索**出现在报错里（若备份还带 officemcp，直接给出
  `OFFICE_PYTHON=<path>`），仍不参与自动选择——备份随时会被清理。
- README 新增《workbuddy 升级后突然 DEGRADED》：三条定位命令，并点明
  「`npm test` 通过 = 降级契约成立，不代表本机环境可用」。
- **无行为变更**：探针逻辑、候选优先级、降级契约均未改。

## 1.2.1

**OCR 审核（阿里 open-code-review，DeepSeek v4-pro）修复批**——17 条发现里的
2 🔴 + 5 🟠 + 6 🟡（office 侧）；knowlp 侧同步见 knowlp-rag 仓 f2a3fd8。

### 修复（office 侧）

- 🔴-1 **二级深扫从未跑通**：source_uri 是 root 相对路径，容器预检与 Workbooks.Open
  却当绝对路径用——每个候选都预检失败、开簿零次。两处补 `os.path.join(root, …)`。
- 🔴-2 DispatchEx 无保护：cands 为空也起 Excel、起失败会把已完成的一级结果全部丢掉。
  现在 `if cands` 守卫 + 自包 try/except，失败降级为 deep[] 一条 error，一级照常返回。
- 🟠-3 magic 判出的 unknown（gzip/rar/7z）进不分类报告（reason: unknown-magic）。
- 🟠-4 未认领扩展名回落 magic 嗅探（改名 .dat 的 PDF 按内容判）；
  无信号文件才落 unrecognized。
- 🟠-5 点文件不再吃 max_files 配额。
- 🟠-6 VBA 探测读到值才 break（16.0 存在但未设值时继续试 15.0/14.0）。
- 🟠-7 内层已带稳定码的异常不再被叠加 KB_SCAN_FAILED 前缀（OPEN_FAILED 可达）。

### 回归钉

- `test/fixtures/make-xlsx.py`：经同一 COM 通道生成**真 xlsx**；
  `test/kb-scan.mjs` 深扫断言取到表头（指标/数值）、used_range、工作表名——
  🔴-1 修复前该用例必红（相对路径 bug 使开簿零次）。**25 断言 PASS，node exit=0。**
- `npm test`：22 工具注册 + schema 契约 PASS。

### P2 追加（夜间窗口②，2026-10-04）

- **P2-a 分类学**（`dd354b4`）：`kb_taxonomy_get` / `kb_taxonomy_set`（工具 #23/#24，
  纯 JS 状态，无 COM）——add / rename / merge（规则并入 + supersedes 记录，evidence 语义）/
  delete（HAS_CHILDREN 保护）；每次 set 自带**漂移审计**（重叠分类 = 同 (kind,pattern)
  出现在两个类别、parent 环、悬空指针），问题随响应返回（audit.clean）。稳定码新增
  DUPLICATE_ID / CATEGORY_NOT_FOUND / HAS_CHILDREN / UNKNOWN_OP。
- **P2-b 分类与影子方案**（本提交）：`kb_classify`（只读——单命中归档、零命中未归类、
  多命中冲突；**冻结桶（not-material / 带 reason）不参与**，红线 9；semantics 如实声明
  unavailable，语义缝属 P3-b）+ `kb_scheme_propose`（只读影子方案——label 链 + 原文件名，
  重名加指纹前 8 位；scheme-<plan_id>.json 落状态目录，不碰被测目录）。
  稳定码新增 NO_TAXONOMY。**工具数 22 → 26。**

### 返修补遗（夜间窗口，2026-10-04）

- `byBucket` 排序键加 `format` 回落（`a.ext || a.format`）——unrecognized 条目里
  unreadable 行只有 `ext`、其余行只有 `format`，缺回落会让同原因码条目排序不稳定。
- 版本落 **1.2.1**（package.json 与 lib VERSION 同改；register.mjs 有版本一致性断言）。
  OCR 修复批当时只写了 1.2.1 段 CHANGELOG、版本号停在 1.2.0——不补则发 npm 会重发 1.2.0。

### ★12 junction 越界修复（2026-10-06 真实复现）

- **红线 3（链接默认不追）在默认模式下是破的**：Windows junction 是重解析点（reparse
  point）不是符号链接——`is_symlink()` 为 False、`os.walk(followlinks=False)` 照样走进去，
  root 内一条 junction 就把扫描带出了 root（真实复现：`link-to-outside/secret.md` 被登记）。
- 修法三层：① subfolder 判据换 reparse（S_ISLNK 或 `st_reparse_tag == IO_REPARSE_TAG_MOUNT_POINT`），
  junction/symlink 目默认不追；② 逐目录 realpath **越界保险丝**——realpath 不在 root 下整棵
  剪掉，与链接类型无关；③ 文件级链接默认不登记、不占配额（OneDrive 等云占位 tag 不在此列，
  照常登记）。knowlp 侧 `pool_scan.py` 同步修（knowlp-rag `6c6a489`，pytest 282）。
- 回归钉：`test/kb-scan.mjs` 新增 junction 用例（`mklink /J` 无需权限，本机真跑）——
  root 外 secret.md 不被登记、root 内文件照常。**27 断言 PASS。**

### 已知未修


- ensureServer 陈旧锁挂起（1.1.0 既有，锁指向死 pid 时无限等待；删锁即恢复）。
- VBA 备用通道仅探测无使用路径（需峄开 AccessVBOM）。
- magic-mismatch 原因码（扩展名与 magic 冲突）未实现。

## 1.2.0

**kb_***：陌生目录的自建分类仓库——第一批（P0+P1，工单 2026-10-03）。目标不是给 vault
加功能，而是让 agent 在陌生环境里自己长出分类体系。身份哈希与 knowlp 侧 pool_scan
逐字同源（sha256(rel_posix\x00size\x00mtime_ns)），词表两仓人工对齐。

### 新增

- `kb_scan_root`（只读，第 22 个工具）：FSO 通道枚举任意目录树（身份字段取 os.stat
  权威 ns 值——FSO 的 DateLastModified 秒精度会让指纹失配，故 FSO 只管枚举）；
  扩展名 + magic 双判据分类；无扩展名/损坏 → `unknown` 显式；容器 → `mixed`；
  不开工作簿的一级扫描为缺省，`max_deep>0` 按配额只读开簿深扫 Excel 结构
  （表头 / UsedRange / LinkSources / 逐表 Hyperlinks）。
- **不分类桶两档**（峄定稿）：`not-material`（本就不是资料：pyc/exe/lnk/url/rev/pack/idx…）
  与 `unrecognized`（认不出、可能有用）分列；每条带原因码
  （no-extension / unclaimed-extension / unreadable）；冻结语义——不进任何下游动作。
- `unclassified.json` 不分类报告：unrecognized 在前、not-material 在后，按原因码+扩展名排序。
- **Obsidian 硬守卫**：`KNOWLP_VAULT` 与 root 任何方向重叠 → `VAULT_FORBIDDEN`，
  纯 JS 先于 COM 执行（降级态也生效），测试钉住。
- 产物：状态目录 `kb/<root_id>/`（registry.json + unclassified.json），绝不写被测目录。

### 真机证据

- `node test/kb-scan.mjs`：守卫三态（==vault / 在 vault 内 / 包含 vault）+ OPEN_FAILED
  全拒绝且不进 COM；合成树 9 文件分类全对（.sample→code、.pyc→not-material、.xyz→
  unrecognized、无扩展名文本→text magic、无扩展名二进制→unknown）；重复扫描身份集全等；
  artifact 只落状态目录、被测目录零 JSON 残留；`vba_channel = VBA_ACCESS_DENIED`
  （本机 AccessVBOM 未开，如实上报即正确行为）；深扫路径假 xlsx 软失败不崩。**18 断言 PASS。**
- `npm test`：22 工具注册 + 只读工具不带 mode 的 schema 契约（kb_scan_root 已纳入）PASS。
- 深扫**容器预检**（PK/OLE2 头）由真机教训驱动：伪造 xlsx 会让 Excel 弹修复对话框挂死
  通道——非法容器现在软失败 `not-a-valid-workbook-container`，绝不递给 Excel。

### 已知未修

- ensureServer 对**陈旧锁文件**（pid 已死）会挂起等待而非自愈——真机实测（锁指向已死
  pid 4316 时新调用无限等待）；删锁即恢复。属 1.1.0 自举层的既有问题，本版未动。
- VBA 备用通道只做了可用性探测，没有实际使用路径（需要时峄手动开 AccessVBOM 后另做）。
- magic-mismatch 原因码（扩展名与 magic 冲突）规格标可选，未实现。

## 1.1.0

三处静默缺陷，全部来自真机复测。**前两处会静默破坏用户数据**；因为改动了两处默认行为，
按语义化版本走次版本而非补丁。

### ⚠ 行为变更

- **`word_edit` / `office_replace_document_terms` 的 `match_case` 默认 `true`（原为 false）。**
  Word 在 `MatchCase=False` 时会把**被找到文本的大小写**套到替换文本上：被替换文本含
  大写拉丁（如 `DSH`）时，替换文本里的 latin 会被强制大写。
  实测读回 run 属性 `ALLCAPS=False` —— 是**真字符被改**，不是显示格式。这违反 Scholarship Lock
  （文字不该被静默改动）。现在默认字面匹配；确需模糊匹配请显式传 `match_case: false`。
- **`office_update_monthly_report` 遇到不合规表头改为报错，不再静默继续。**
  原先判据只看「期间」列，于是「期间/科目/金额」这种表被判成自家产物：清空后按
  `MONTHLY_HDR` 重写，`_row.get("借方")` 取不到 → **金额整列变 null，却仍报 `kept_rows=3` 成功**。
  现在判据覆盖全部必需列（期间/借方/贷方），缺任一列即拒，报 `[SCHEMA_MISMATCH]` 并列出
  缺失列与实有列。**原先"成功"的调用现在会失败 —— 这是有意的。**

### 修复

- **`word_edit` / `office_replace_document_terms`：计数与实际替换口径不一致。**
  计数那遍用关键字参数、没设 `MatchCase`，走的是 Word Find 的**残留状态**；两遍可能一个
  区分大小写、一个不区分，表现为"数出 N 处、实际替换 M 处"。现在两遍同口径。
- **按名字取工作表失败时不再抛裸 COM 异常。**
  `(-2147352567, '发生意外。', (0, None, None, None, 0, -2147352565), None)` 是
  `DISP_E_EXCEPTION` 包着 `DISP_E_BADINDEX`，从字面看不出"工作表不存在"，排查成本极高。
  新增 `PREAMBLE` 里的 `_sheet()` helper，报 `[SHEET_NOT_FOUND]` **并列出实有表名**；
  8 处未受保护的取表点改用它。

### 已知未修

- `excel_journal_post` / `office_generate_accounting_report` 写入的日期形如
  `2026-09-01T00:00:00+00:00` 而非 `2026-09-01`（观感/后续处理问题，不影响计算）
- `excel_read_range` 对公式单元格可能返回 null（要拿计算值须配合 `excel_recalc`）

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
