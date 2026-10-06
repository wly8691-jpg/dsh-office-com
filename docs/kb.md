# kb_* — 陌生目录的自建分类仓库（v1.2.1）

- 依据：《工单-陌生目录自建分类仓库-Office×KnowLP-CC-20261003》（P0–P3 交付）＋《工单-Office资料写回-CC-20261006》《工单-Office属性落盘根因-B4-CC-20261006》
- 定位：让 agent 在**陌生环境**里靠这族工具登记一棵目录树的原生资料、自己长出分类体系——**不是**给 Obsidian vault 加功能。

## 能力（v1.2.1 = P0–P3 交付，8 个工具）

| 工具 | 读/写 | 说明 |
|---|---|---|
| `kb_scan_root` | 只读 | 登记任意目录树：扩展名 + magic 双判据分类；认不出 → 「不分类」桶（两档 + 原因码）；`max_deep>0` 按配额只读开簿深扫 Excel 结构；产物落状态目录 |
| `kb_taxonomy_get` / `kb_taxonomy_set` | 只读 / 变更 | 分类学（自长分类）：增/改名/合并/删 + 漂移审计（重叠/环/悬空），落 `kb/<root_id>/taxonomy.json` |
| `kb_classify` | 只读 | 机械规则分类（ext / name-contains / path-contains）；冻结桶不参与；`use_semantics=true` 走 KnowLP 语义缝（见下） |
| `kb_scheme_propose` | 只读 | 影子目录方案（label 链 + 原文件名，重名加指纹前 8 位）；`scheme-<plan_id>.json` 落状态目录，**不碰任何文件** |
| `kb_annotate` | 变更 | 写标注：`annotations.jsonl` 每行是 ContextItem 形状 + 标注字段（登记优先）。文件已不在/指纹已变 → 如实标 `unverifiable`（`require_intact` 可改为直接拒绝）。写回原文件默认关——`write_back && confirm` 双闸 |
| `kb_verify` | 只读 | 核账：标注可解析 / 源文件仍在且指纹吻合 / 影子方案与磁盘一致；问题清单随响应返回（`clean`）。`office_properties:true` 增一路只读读回：逐个打开 Office 标注文件读回 `kb_*` 属性与登记比对（`property-mismatch` / `property-missing`，绝不猜） |
| `kb_write_back_office` | 变更 | 对单个 Office 文件写/删 `kb_*` 自定义文档属性——**只碰属性面板，绝不碰单元格/正文/幻灯片**；`preview` 只读打开报告现值与将写值；写后自动关掉重开核账（属性读回一致 + 内容快照 sha 一致）；`rollback` 删同名属性并同样核账 |

## 组合面：KnowLP 语义缝

**这一族就是 Office × KnowLP 的组合面**——Office 知道「这是什么文件、能不能真被打开」，KnowLP 知道「它在讲什么」。

`kb_classify(use_semantics=true)` 把待归类的资料交给 KnowLP 检索，走**三级梯子**，每一级都如实回报走到了哪：

| 级 | 含义 |
|---|---|
| `approved` | 经 `ctx.tools.execute` 审批派发（正路） |
| `direct` | 拿不到审批通道时降级直连（仅降级用） |
| `unavailable` | KnowLP 不可用 → 如实声明 + `next_call_suggestion` |

两条踩出来的细节：

- **工具名按注册表里的实际名字解析**，不写死。MCP 接入下注册名是 `mcp__<server>__*` 形状，而 `get()` 是精确查表——写死 `knowlp_search` 必然查不到，异常还会被静默吞成「0 命中」。
- **断链 / 失败 / 不可用一律显式降级并带回错误**（`approved-failed` / `direct-failed` + `seam_errors` 入 output），**绝不假装做了语义判断**。

### 身份与词表对齐纪律（跨仓，必须两边一起改）

- **身份哈希**：`sha256(rel_posix \0 size \0 mtime_ns)` —— 与 knowlp 侧 `pool_scan.fingerprint_of` **逐字同源**，两边都不许另立公式。mtime 取 `os.stat` 权威 ns 值（FSO 的 DateLastModified 只有秒精度，只用于枚举，不进指纹）。
- **词表**：本仓 `KB_POOL_EXTENSIONS` / `KB_NOT_MATERIAL` ↔ knowlp 侧 `pool_scan.POOL_EXTENSIONS` / `NOT_MATERIAL_EXTENSIONS`。任何一边改词表，**同一提交里改另一边**（各自测试同步）。

## 产物位置

- `<DSH_OFFICE_STATE_DIR | ~/.dsh-office-com>/kb/<root_id>/`
  - `registry.json` —— 全量登记（含 stats / deep / vba_channel）
  - `unclassified.json` —— **不分类报告（交付物，非副产品）**：`unrecognized` 在前（唯一需要人眼的池）、`not_material` 在后（登记即可），按原因码 + 扩展名排序
  - `taxonomy.json` —— 分类学；`scheme-<plan_id>.json` —— 影子方案
  - `annotations.jsonl` —— 标注。**追加式**：重标注会新增一行，取最新一条
- `root_id = sha256(realpath(root))[:16]`（沿用 1.0.2 约定）。**绝不写进被测目录。**

## 不分类桶（P1.5，峄定稿）

**一句话规则：认不出来的一律进「不分类」，显式登记、绝不猜、绝不处理。**

| 档 | 含义 | 现阶段判据 |
|---|---|---|
| `not-material` | 本就不是资料（派生物/可执行/快捷方式/VCS 内部件） | 扩展名 ∈ KB_NOT_MATERIAL（`.pyc .exe .dll .lnk .url .rev .pack .idx .msi .class`） |
| `unrecognized` | 认不出、可能有用 | 原因码：`no-extension`（无扩展名且 magic 不匹配）/ `unclaimed-extension`（扩展名没人认领）/ `unreadable` |

- **冻结语义（命门）**：不分类条目排除在任何下游动作之外——不标注、不进影子方案、不清理、不进分类流（kb_classify/kb_scheme_propose 均跳过）。要动它，必须先把目标类别**显式确认**；默认永不自动处理。
- `.sample` / `.example`（实测 31 个）已并入 **code** 池，不落不分类。
- 容器（`.zip .7z .rar .tar .gz`）**不猜**：无扩展名时 magic 判 `mixed`；带扩展名但无池认领 → **回落 magic 嗅探**（🟠-4，2026-10-03：改名 .dat 的 PDF 按内容判；有信号归池，无信号才落 unrecognized）；magic 认出但无池可归（gzip/rar/7z）→ unrecognized + 原因码 `unknown-magic`（🟠-3）。

## 边界与规则（显式声明，不许默默如此）

1. **点文件/点目录整类跳过**（`.env` / `.gitignore` / `.obsidian` / `.git` …）——`.env` 这类文件**连登记都没有**，这是继承 knowlp 侧 pool_scan 的既有行为，**保留**；需要登记它们时必须显式改规则并在此记录。
2. **标注指纹语义（★4/★5，2026-10-05）**：`kb_annotate` 成功写回后会**重新 stat 并把写回后的新指纹存进标注**——`kb_verify` 直接对新指纹核账，**不再报警**；但 `registry.json` 里的旧指纹要**重扫（kb_scan_root）才刷新**。写回失败/未写回时，标注存调用方给的指纹，文件若已变 → verify 如实报 `fingerprint-changed`。
3. **sensitivity 缺省 = private（★7，2026-10-05 峄定）**：标注未显式给 sensitivity 时落 `private`——**缺省拒绝云回退**，要放行某类资料需显式打 `public`（本 vault 过半为内部/商业件，缺省 public = 默认外送，是有害默认）。
4. **写回载体 = 自定义文档属性（2026-10-06 第九批起）**：写回分两路——`.md` 走**前置元数据**；Office 文件（xlsx/xlsm/xls、docx/doc、pptx/ppt）走 **CustomDocumentProperties**，**只写 `kb_*` 键、绝不碰内容**，且 `kb_annotate` 的 Office 写回**只走审批通道**（拿不到审批就不写，只登记）。双闸 `write_back:true && confirm:true`。
   - ⚠️ **`write_back: true` 表示写回操作成功，不等于文件字节变了**——落盘与否看 `output.disk_changed`。
   - ⚠️ **三种格式行为不同（B4 修复的副作用）**：**Word / PowerPoint 上「写同值」也会真的重写文件**（根因是它们不把自定义属性变更视为文档修改，`Save()` 会静默什么都不写；修复要求保存前清 `Saved` 标志，等于强制落盘一次），所以「同值写回幂等、无副作用」在 word/ppt 上**不成立**；**Excel 会合理跳过重写**。三种格式的 `disk_changed` 都如实回报，**别按同一条直觉推断**。
5. **Obsidian 硬守卫**：`KNOWLP_VAULT`（env）解析出的真实路径与 root **任何方向重叠**（相等/互相包含）→ `VAULT_FORBIDDEN`，纯 JS 判断、先于 COM 通道（降级态也生效）。有测试钉住。
6. **符号链接 / junction 默认不追（★12，2026-10-06 三层修）**：① 子目录判据换成 reparse 点（`S_ISLNK` 或 `st_reparse_tag == IO_REPARSE_TAG_MOUNT_POINT`）——Windows junction 是重解析点不是符号链接，`is_symlink()` 恒 False；② **逐目录 realpath 越界保险丝**：realpath 不在 root 下整棵剪掉，与链接类型无关；③ 文件级链接默认不登记、不占配额（OneDrive 等云占位 tag 不在此列，照常登记）。已见目录（realpath 去重）防环。
7. **v1 无任何 move/rename/delete 代码路径**——影子方案只出方案不动文件；`kb_apply_scheme`、journal、undo 整体后置。
8. **VBA 备用通道**：主路是 COM 对象模型（FSO + Excel 只读开簿）。VBA 通道需要 Excel 信任中心「信任对 VBA 工程对象模型的访问」（`AccessVBOM`）——本机未开启，工具如实报 `VBA_ACCESS_DENIED`；不假装成功、不静默跳过、不写进安装步骤。需要时由峄手动开。
9. **深扫容器预检**：`.xlsx/.xlsm` 必须 PK 头、`.xls` 必须 OLE2 头——非法容器（截断/伪造件）直接软失败 `not-a-valid-workbook-container`，**绝不递给 Excel**（真机实测：伪 xlsx 会触发修复对话框挂死通道）。

## 身份命名空间契约（OCR 🟠-8，2026-10-03）

- 扫描身份（fingerprint / source_uri）**按 root 作用域**：root 相对路径 + size + mtime_ns。
- 跨 root **不得合并**裸 entries 字典（同相对路径+同大小+同 mtime 会互撞）——kb 工具按
  `root_id = sha256(realpath(root))[:16]` 分目录存放，天然分域。

## 明确没做什么（本版本）

- 影子方案的 **apply / journal / undo**；`kb_clean` / `kb_dedupe` —— 现在只到「出方案 + 核账」
- **VBA 通道的实际使用**（只做了可用性探测与如实上报）
- **magic-mismatch 原因码**（扩展名与 magic 冲突的检测，规格标可选，未实现）
- **跨 root 的合并视图**（按设计分域，不是待办）
- Office 写回**不碰内容**——正文/单元格/幻灯片一律不动，这是设计边界不是缺项

---

（实现：CC 2026-10-03 起，P0–P3 加 Office 写回两批，最新 2026-10-07。测试：
`npm test` → `test/register.mjs`（29 工具注册 + schema 契约，**不需要 Office**）。kb 相关：
`test/kb-scan.mjs`（守卫三态 / 分类两档 / 指纹稳定 / artifact 隔离 / VBA 如实上报 / 深扫不崩 / junction 越界）、
`test/kb-taxonomy.mjs`（增改合并删 + 漂移审计 + 降级梯子）、`test/kb-writeback.mjs`、
`test/kb-writeback-office.mjs`（真 Office）、`test/kb-writeback-routing.mjs`、
`test/kb-nested-dispatch.mjs`（B2 + §1.4 钉，不需要 Office）、`test/kb-result-contract.mjs`（两条纯判据，不需要 Office）、
`test/faults.mjs`（故障注入）、`test/leakcheck.mjs`（深扫后零残留 Office 进程）、
`test/kb-corpora.mjs`（四语料实测，需真机 Office）。）
