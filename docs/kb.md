# kb_* — 陌生目录的自建分类仓库（v1.2.0）

- 依据：《工单-陌生目录自建分类仓库-Office×KnowLP-CC-20261003》（P1 交付）
- 定位：让 agent 在**陌生环境**里靠 `kb_scan_root` 登记一个目录树的原生资料、自己长出分类体系——**不是**给 Obsidian vault 加功能。

## 能力（v1.2.0 = P0+P1 交付）

| 工具 | 读/写 | 说明 |
|---|---|---|
| `kb_scan_root` | 只读 | 登记任意目录树：扩展名 + magic 双判据分类；认不出 → 「不分类」桶（两档 + 原因码）；`max_deep>0` 按配额只读开簿深扫 Excel 结构；产物落状态目录 |

### 身份与词表对齐纪律（跨仓，必须两边一起改）

- **身份哈希**：`sha256(rel_posix \0 size \0 mtime_ns)` —— 与 knowlp 侧 `pool_scan.fingerprint_of` **逐字同源**，两边都不许另立公式。mtime 取 `os.stat` 权威 ns 值（FSO 的 DateLastModified 只有秒精度，只用于枚举，不进指纹）。
- **词表**：本仓 `KB_POOL_EXTENSIONS` / `KB_NOT_MATERIAL` ↔ knowlp 侧 `pool_scan.POOL_EXTENSIONS` / `NOT_MATERIAL_EXTENSIONS`。任何一边改词表，**同一提交里改另一边**（各自测试同步）。

## 产物位置

- `<DSH_OFFICE_STATE_DIR | ~/.dsh-office-com>/kb/<root_id>/`
  - `registry.json` —— 全量登记（含 stats / deep / vba_channel）
  - `unclassified.json` —— **不分类报告（交付物，非副产品）**：`unrecognized` 在前（唯一需要人眼的池）、`not_material` 在后（登记即可），按原因码 + 扩展名排序
- `root_id = sha256(realpath(root))[:16]`（沿用 1.0.2 约定）。**绝不写进被测目录。**

## 不分类桶（P1.5，峄定稿）

**一句话规则：认不出来的一律进「不分类」，显式登记、绝不猜、绝不处理。**

| 档 | 含义 | 现阶段判据 |
|---|---|---|
| `not-material` | 本就不是资料（派生物/可执行/快捷方式/VCS 内部件） | 扩展名 ∈ KB_NOT_MATERIAL（`.pyc .exe .dll .lnk .url .rev .pack .idx .msi .class`） |
| `unrecognized` | 认不出、可能有用 | 原因码：`no-extension`（无扩展名且 magic 不匹配）/ `unclaimed-extension`（扩展名没人认领）/ `unreadable` |

- **冻结语义（命门）**：不分类条目排除在任何下游动作之外——不标注、不进影子方案、不清理、不进 ContextItem 流。要动它，必须先把目标类别**显式确认**；默认永不自动处理。
- `.sample` / `.example`（实测 31 个）已并入 **code** 池，不落不分类。
- 容器（`.zip .7z .rar .tar .gz`）**不猜**：无扩展名时 magic 判 `mixed`，带扩展名时 unclaimed → unrecognized。

## 边界与规则（显式声明，不许默默如此）

1. **点文件/点目录整类跳过**（`.env` / `.gitignore` / `.obsidian` / `.git` …）——`.env` 这类文件**连登记都没有**，这是继承 knowlp 侧 pool_scan 的既有行为，**保留**；需要登记它们时必须显式改规则并在此记录。
2. **Obsidian 硬守卫**：`KNOWLP_VAULT`（env）解析出的真实路径与 root **任何方向重叠**（相等/互相包含）→ `VAULT_FORBIDDEN`，纯 JS 判断、先于 COM 通道（降级态也生效）。有测试钉住。
3. **符号链接/junction 默认不追**；已见目录（realpath 去重）防环。
4. **v1 无任何 move/rename/delete 代码路径**——影子方案（P2）只出方案不动文件；`kb_apply_scheme`、journal、undo 整体后置。
5. **VBA 备用通道**：主路是 COM 对象模型（FSO + Excel 只读开簿）。VBA 通道需要 Excel 信任中心「信任对 VBA 工程对象模型的访问」（`AccessVBOM`）——本机未开启，工具如实报 `VBA_ACCESS_DENIED`；不假装成功、不静默跳过、不写进安装步骤。需要时由峄手动开。
6. **深扫容器预检**：`.xlsx/.xlsm` 必须 PK 头、`.xls` 必须 OLE2 头——非法容器（截断/伪造件）直接软失败 `not-a-valid-workbook-container`，**绝不递给 Excel**（真机实测：伪 xlsx 会触发修复对话框挂死通道）。

## 明确没做什么（本版本）

- `kb_taxonomy_get/set`、`kb_classify`、`kb_scheme_propose`（P2）
- `kb_annotate`、`kb_verify`、knowlp 语义缝与降级梯子（P3）
- 影子方案的 apply / journal / undo；`kb_clean` / `kb_dedupe`
- VBA 通道的实际使用（只做了可用性探测与如实上报）
- magic-mismatch 原因码（扩展名与 magic 冲突的检测，规格标可选，未实现）

---
（实现：CC 2026-10-03。测试：`node test/kb-scan.mjs`（守卫三态 + OPEN_FAILED + 分类/两档/指纹稳定/artifact 隔离/VBA 如实上报/深扫不崩）；`npm test`（22 工具注册 + 只读无 mode schema 契约）。）
