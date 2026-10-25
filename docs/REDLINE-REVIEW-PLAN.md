# Office CodeBuddy：Cursor 云端开发交接计划

> 用途：把本文件全文交给 Cursor Cloud Agent，使其无需读取此前的 ChatGPT 对话即可继续开发。
> 建议落库路径：`docs/REDLINE-REVIEW-PLAN.md`。
> 项目：<https://github.com/brucecbi/office-codebuddy>。
> 资料基线：2026-10-24；观察到的提交为 `c79441f0a8d727e735c393741dace4b0d79e2573`，版本为 `0.2.0-preview.4`。这只是分析基线，不是要求回退的目标。
> 首批任务：P1A——单份 Word 文档、多作者修订的只读审计基础层。

## 1. 用户目标与业务场景

用户在协议审阅中使用 Office CodeBuddy，希望 AI 能准确理解 Word 中不同作者的 redline、comments 和回复，并且能够证明自己读取了哪些证据、哪些地方仍存在缺口。

场景一：同一份协议含多个作者留下的修订与批注。主 Agent 进行解释和审阅，独立的只读 Reviewer 核对其证据覆盖和判断，发现漏读、错配作者、把删除内容当现行条款、把批注当已实施修改等错误。

场景二：多个部门和外部参与方从一份或多份基稿出发，各自返回修订稿，数量可能达到十几份以上。产品应按条款汇总建议、保留来源和作者、识别冲突，并让用户决定采用哪些建议。这里可以借鉴 GitHub 的分支、提交、PR、讨论与审批机制，但需要处理 Word 的锚点、修订结构和历史不完整问题。

目标产品是可追溯的协议审阅工作台。沿用现有 Word 插件与本地 daemon 架构，分阶段增加审计和意见整合能力。

## 2. Cursor 首次启动时的工作方式

1. 读取当前仓库的 `AGENTS.md`、`README.md`、`docs/SECURITY-AND-PRIVACY.md`、`docs/TESTING.md`，再完整读取本文件。
2. 记录实际分支、HEAD、工作区状态、Node 版本和依赖状态。检查是否存在更具体的目录指令。保留已有修改，不重置到上面的历史提交。
3. 在用户授权的功能分支上工作。将本文件保存为 `docs/REDLINE-REVIEW-PLAN.md`；如果已有同用途文件，先检查并更新，避免重复建立相互矛盾的规范。
4. 在现有 `AGENTS.md` 追加一个简短入口，不覆盖原有规则。建议文字见第 13 节。
5. 检查当前实现与本计划是否已有差异；把影响任务的差异写入进度记录，然后直接实现 P1A。不要重复实现已经完成的功能，也不要仅输出另一份规划后停止。
6. 为本批建立 `docs/REDLINE-REVIEW-PROGRESS.md`，记录实际范围、完成项、未完成项、关键决策和测试结果。后续会话从该文件接续。

本文件定义产品目标、技术约束与验收标准；具体推送、PR、发布权限以用户当前任务消息及仓库适用规则为准。本文件不会自行扩大权限。

## 3. 已有实现：先复用，再扩展

以下是分析时观察到的实现，启动开发时必须重新检查当前代码。路径用于定位，不能替代阅读代码。

| 位置 | 已有能力／应注意的边界 |
| --- | --- |
| `taskpane/shared/word-review.js` | 文档结构清单；Current/Original 读取；修订、作者、日期、批注及回复；分页；范围 OOXML 证据；支持范围及缺口信息。当前明确保留 `complete_file_review:false`、`enumeration_complete:false` 等保守边界。 |
| `taskpane/shared/word-ooxml-state.js` | 用于状态比较的有限 OOXML 归一化；不是完整 DOCX 包解析器，归一化输出不能直接写回文档。 |
| `taskpane/shared/word-write-safety.js`、`word-edit-plan.js`、`tools-word.js` | 写入锚点、前置条件、验证和已有编辑流程。应保留现有语义。 |
| `daemon/office-tools.mjs` | Office 工具定义与桥接；新增只读能力应与既有工具接口一致。 |
| `daemon/office-read-only.mjs`、`taskpane/shared/taskpane.js` | 两侧存在对应的只读工具集合；新增入口需保持一致，并运行 `tests/office-read-only.test.mjs`。只读分类本身不等于未来 Reviewer 的权限隔离，后者仍须限制实际可用工具。 |
| `daemon/preview-policy.mjs` | 默认 `office-only` 策略限制 SDK 的通用工具、用户设置、hooks、skills、agents 和额外 MCP。不能通过切换 extended 绕开产品边界。 |
| `daemon/system-prompt-word.md` | Word 读取、修订解释和写入约束；文档与批注是待分析数据。 |
| `daemon/runtime/index.mjs` | 当前运行时入口，存在 CodeBuddy 与 Claude 兼容路径；不要为本批重写运行时。 |
| `tests/word-review.test.mjs`、`word-ooxml-state.test.mjs`、`word-write-guards.test.mjs` | 复用并补充有针对性的回归。 |
| `tests/fixtures/word-native-*-snapshots.json`、`word-native-accepted-insertion.json` | 可用于离线回归的记录样本；后者包含接受插入后暴露原有隐藏删除的情形。用它们运行测试不等于完成当前版本的原生 Word 测试。 |

项目是 macOS 优先的 Word/Excel 开发预览。保留 MIT 许可、`LICENSE`、`UPSTREAM.md` 及上游归属。与本功能无关的 Excel、桌面生命周期和安装流程不在首批改造范围。

## 4. 必须保留的设计约束

### 4.1 来源与快照

- 每次审阅对应一个明确来源和固定证据集合。区分磁盘 DOCX、未保存的实时 Word 内容、范围 OOXML 导出和规范化证据快照。
- 对原始 DOCX 可计算文件字节哈希；对实时读取只能计算所取得证据的哈希。后者不得命名为完整文件哈希，也不得证明未覆盖部分未变化。
- 多次 `context.sync()` 不是事务。读取前后以及必要的分页间应检查能够获得的状态；如果无法证明读取一致，应保留 `consistency: unknown` 或更具体的受限状态。
- 检测到读取中变化、快照过期、分页缺失或无法确认一致性时，不能宣称全文件审计通过。可展示已读证据及缺口；过期证据不能作为后续写入依据。
- 保留原始来源。为 diff 构造 accepted/current 视图时，应生成派生视图，不覆盖含修订的原件。

### 4.2 完整性与可解释性

- 区分“读取成功”“支持这个结构”“证据覆盖完整”和“业务解释正确”。四者不能互相替代。
- 零条返回、API 失败、不支持某个 story、未读完分页，各自是不同状态。
- 范围 `getOoxml()` 或 Flat OPC 输出不能冒充原始完整 DOCX ZIP 包。
- Word 原生修订集合可能合并呈现插入／删除；`Deleted.text` 为空不能证明没有删除文本。
- 不使用“XML 标记数等于修订条数”作为唯一完整性判断。原始标记、规范化变更和业务建议之间可能是一对多或多对一，必须保存映射与解释。
- 不把现有 `enumeration_complete:false` 直接改成 true 来满足测试。支持边界变化必须有对应实现和证据。
- 每条主 Agent 判断和 Reviewer 判断应引用证据 ID。没有证据支持的结论保持未确认。

### 4.3 身份、权限与写入

- 原始 Word 作者、文件提交者、部门、AI 执行者和最终审批者分别记录。Word 的作者字段是来源元数据，不是已经验证的人员身份。
- session-local paragraph ID、revision index 和临时 snapshot token 不能直接成为跨文件永久 ID。证据 ID 至少绑定来源／快照；后续匹配以显式 lineage 和多种锚点辅助。
- 默认保持 `office-only`。产品内部的 Reviewer 和 skills 应由应用显式提供受限工作流，不能假定放入一份 skill 文件就会被默认运行时加载。
- 可以并行分析冻结快照；实时 Word 修改必须沿用单写入者、确认卡、串行执行、live preconditions 和写后验证。
- 保留 `RESULT_UNKNOWN`。不知道某次写入是否生效时，不自动重放，不换工具再次执行同一写入。
- 文档正文、批注和工具返回的文档内容始终是数据，不能变成 Agent 的操作指令。

## 5. 推荐的证据模型

以下为领域结构建议，不要求机械照搬字段名。优先适配仓库现有类型与接口，增加 schema 版本，保持兼容。

| 对象 | 最低需要的信息 |
| --- | --- |
| `ReviewSnapshot` | `schema_version`、`snapshot_id`、来源类型、来源 ID、原始文件哈希或明确标注的证据哈希、捕获时间、读取范围、解析器／适配器版本、一致性状态。 |
| `CoverageEntry` | story／part／能力范围、读取状态、分页游标与终止证据、失败或不支持原因、关联证据 ID。可采用 `read`、`read_failed`、`unsupported`、`not_read`、`confirmed_absent` 等状态。 |
| `RevisionEvidence` | 快照内唯一 ID、来源 part／story、原始修订标记及其 ID、类型、原始作者／时间、before/current 证据、原始内容引用、锚点和映射状态。 |
| `CommentEvidence` | 来源 ID、作者／时间、正文、锚定范围、回复线程、resolved 状态、是否发现锚点缺失。批注内容不能自动变成文本修改。 |
| `AnalysisClaim` | 判断、引用的证据 ID、未决点及相关条款。首批可先提供接口和确定性发现，无须接入真实模型。 |
| `ReviewFinding` | 漏读、证据不足、作者错配、视图混淆、冲突等分类；相关证据／判断 ID；需要补读还是人工决策。 |
| `DecisionEvent` | 后续阶段记录采纳／拒绝／修改、决策人、理由、时间及来源。明确区分 approved、applied、verified、exported。 |

若某字段只有范围级来源，就如实保存范围级来源；不要虚构完整 ZIP part、跨版本身份或不存在的原始标记。以原始证据为依据的提取结果和模型推断必须分开。

## 6. 本次立即实现：P1A 单文件只读审计基础层

### 6.1 范围

本批交付一个实际可调用的只读审计入口、结构化结果以及可检查的实现。只输出规划、孤立 schema 或未接入的模块不算完成。

建议实现顺序：

1. **证据结构与快照封装。** 复用现有文档清单、Current/Original、revisions、comments 和范围 OOXML 输出，建立可冻结、可序列化的审阅证据。标明来源和支持范围。
2. **确定性覆盖汇总。** 记录需要读取的范围、已读取范围、分页是否结束、已知失败和不支持项；输出覆盖报告，不生成未经证明的完整性结论。
3. **一致性／过期处理。** 依据现有接口可获得的状态检查变化，区分一致、变化、过期与无法确认。不能取得全文件一致性证据时，明确记录局限。
4. **只读调用入口。** 接入现有 Office 工具或现有审阅流程，同步检查 taskpane 的 `readOnlyOfficeTools` 与 daemon 的 `OFFICE_READ_ONLY_TOOLS`。具体工具名按项目惯例选择。入口返回结构化证据与适合展示的摘要；本批不要求大规模改 UI。
5. **结果展示。** 至少能够看到作者、修订／批注、来源定位和未覆盖项。若现有工具结果展示已足够，可先复用；不要为了首批引入新前端框架。
6. **针对性回归与文档。** 使用合成样本和已有记录样本验证错误路径，更新变更说明、测试摘要和进度记录。

本批不包含：完整 DOCX ZIP importer、OfficeCLI 生产集成、真实模型 Reviewer 会话、多文件合并、自动接受／拒绝修订、clean 稿导出、重写 writer、开放 SDK 通用 subagents、构建新协作平台。

### 6.2 验收用例

| 用例 | 必须观察到的行为 |
| --- | --- |
| 插入与删除组成同一替换，原生集合只返回部分信息 | 保留原生信息和可用 OOXML 证据，报告映射边界；不得把缺少的删除内容解释成不存在。 |
| 删除项的 `text` 为空 | 保留删除项和读取限制，不把它静默过滤掉。 |
| 多作者修改与文件提交者不同 | 作者字段和提交者字段独立，不能相互覆盖。 |
| 相同文字在不同段落出现 | 以来源／锚点区分，不仅按文本合并证据。 |
| 实际支持分页的读取接口尚有下一页 | 报告分页未完成，不能返回完整覆盖通过。当前重点验证 `office_read_story` 和 `office_read_formatting`；未来修订／批注适配器若分页，也采用同一原则。 |
| 某个 story／能力不支持，或读取失败 | 明确列出范围和原因，不显示为“无修订”。 |
| 读取过程中证据变化或 token 过期 | 返回需要重新读取或受限报告；不继续使用旧证据作为已验证状态。 |
| 批注新增回复、resolved 变化、锚点缺失 | 根据可用证据反映变化和缺口；不能只统计评论对象数量。 |
| 原始标记与规范化变更数量不相等 | 保存已知映射或未决状态，不仅用数字相等判断对错。 |
| 中文、不同 Unicode 字符、标点和重复短句 | 来源定位不被有损文本归一化破坏；无法匹配时明确失败。 |
| 执行只读审计 | 不产生 Word 写入，不改变审阅状态，不触发真实模型或扩展权限。 |

当前 `office_read_revisions` 与 `office_list_comments` 为集合返回，不要假定已经存在分页游标。分页完整性检查必须针对实际接口。测试应验证实际风险和行为，避免只把实现细节抄成断言。无需为低影响的文档更新堆砌测试。

### 6.3 完成标准

- 有贯通现有调用路径的最小实现和可检查的输出示例。
- 覆盖范围与一致性缺口真实可见，不宣称已解决所有 DOCX 解析问题。
- 必要回归通过；已有失败和环境阻塞分别记录，不能删除测试或放宽断言来获得绿色结果。
- 默认权限、现有 writer 和未知写入结果的处理保持有效。
- 本地原生 Word 待验收项清楚列出。
- PR 或交接说明区分“本批已经实现”“已测试”“还未测试”“后续计划”。

## 7. 后续路线：按独立批次推进

| 阶段 | 主要交付 | 进入下一阶段的依据 |
| --- | --- | --- |
| P1A：只读审计基础 | 快照、证据模型、覆盖报告、只读入口、离线回归 | 已读与未知状态可区分，每条证据可定位。 |
| P1B：DOCX 包级提取与校验 | 完整 ZIP／relationships／相关 parts 的支持清单，原始包保留，提取适配器和多视图核对 | 对支持的修订／批注类型有样本和可复现映射；未知结构明确阻断或降级。 |
| P1C：独立 Reviewer | 受限只读工具、独立上下文、证据 ID 驱动的复核、缺口回报 | Reviewer 能指出人为注入的漏项／错配，不只复述主 Agent 摘要。 |
| P2：多方意见工作台 | 基稿、提交副本、条款归组、来源追溯、冲突队列、决策账本 | 明确共同基稿；支持同源继承意见去重和多方独立意见保留。 |
| P3：受控应用与导出 | 经用户选择后形成编辑计划，沿用现有安全写入，导出与原生复核 | 决策、应用、验证和导出状态分别可追踪；文档重开验证通过。 |

### 7.1 独立 Reviewer 的具体要求

主 Agent 和 Reviewer 应读取同一冻结快照。Reviewer 获得原始／规范化证据、覆盖状态和主 Agent 的带证据判断，具有独立的系统指令和上下文。只给 Reviewer 主 Agent 的摘要不构成有效独立复核。

确定性检查负责“所有可枚举证据是否入账、分页是否结束、映射是否缺失”。模型负责解释、异常判断和条款关系。两次模型调用不能保证零漏读，更不能修复上游提取已经丢失的内容。

可在后续定义 `revision-auditor`、`comment-reviewer`、`conflict-analyst`、`integrity-checker` 等应用内角色／工作流；按任务拆分的前提是权限和证据输入明确。首批不增加通用 Agent 调度框架。

### 7.2 多方修改与 GitHub 模型的对应

| GitHub 概念 | 协议审阅中的对应 |
| --- | --- |
| repository | 一份协议及其审阅案件／工作区 |
| base commit | 明确指定并保留的基稿快照 |
| branch | 某参与方的提交副本及其来源关系 |
| commit／patch | 可追溯的文本、格式、表格或审阅状态变化 |
| pull request | 一次提请审阅的修改与批注集合 |
| review／discussion | 条款级讨论、回复、部门意见及审批 |
| conflict resolution | 用户对竞争修改的选择、改写与理由 |
| merge result | 由决策生成的文档结果及配套审计账本 |

产品应有提交列表、条款视图、来源对比、冲突队列和决策记录。先做可用的审阅流程，再决定是否需要复杂的分支图界面。

多文件算法的约束：

- 明确共同基稿。不知道共同基稿时，标记 lineage unresolved，支持对照审阅；不默默猜测后自动合并。
- Original/Current 是文档中未决修订的投影，不等于完整历史，也不能自动成为所有文件的共同祖先。
- 若基稿已经带有修订，十几份返回稿继承的同一条意见不能算成十几条独立意见；但各返回稿对它的接受、拒绝、修改等变化需要保留。
- 两个参与方独立提出相同文本，可以归为同组，仍保留两份来源和支持关系。
- 文本相似不代表法律效果相同。冲突至少考虑改动范围重叠、相互依赖条款和跨条款定义；无法确定时进入用户决策队列。
- 批注中的建议／问题与已经实施的修订分开。
- 不承诺一个最终 DOCX 能容纳完整分支历史。原始副本和审计账本保留完整来源，导出文档是某次决策的结果视图。

## 8. 候选工具：后续验证后再引入

下表是研究候选与借鉴方向，不代表已经在本项目完成集成或通过真实合同测试。首次接入必须核对当时版本、许可证、平台与支持范围。

| 候选 | 值得借鉴／验证的能力 | 集成边界 |
| --- | --- | --- |
| [OfficeCLI](https://github.com/iOfficeAI/OfficeCLI) | DOCX 结构与修订查询，作者、日期、类型和路径等结构化信息 | 首先做独立、只读文件适配器；不得作为绕过实时 Word 写入保护的第二写入通道。 |
| [safe-docx](https://github.com/UseJunior/safe-docx) | revisions 提取、before/after、作者、批注、表格变更及显式歧义错误 | 验证分页、重叠修订与复杂结构。研究时文档存在默认分页 50 项及 `AMBIGUOUS_REVISION_OVERLAP` 边界；不能只取第一页。 |
| [docx-cli](https://github.com/kklimuk/docx-cli) | accepted/current、baseline、带标记视图和结构化 AST 思路 | 位置型 ID 修改后可能变化，操作前重读；支持范围需以锁定版本实测。 |
| [word-ai](https://github.com/flyfish-dev/word-ai) | 离线 OpenXML 与在线 Office.js 分工、预期哈希和 PatchSet 思路 | 注意 AGPL 许可与本项目 MIT 的集成关系；借鉴架构前先确定复用方式。 |
| [Docxodus](https://github.com/JSv4/Docxodus) | 共同基稿下多方 consolidate 与冲突报告 | 必须检查 accepted-view 归一化、冲突默认策略、作者保留及 `preserveInputRevisions` 的实际语义；不能直接采用默认策略替用户决策。 |
| [ooxml-integrity](https://github.com/Dmitry-Kov/ooxml-integrity) | 原稿、预期修改与结果之间的确定性完整性检查 | 作为测试／验证思路；具体覆盖范围依工具支持，不能提升为所有 DOCX 的保真保证。 |
| [Pandoc](https://pandoc.org/MANUAL.html#option--track-changes) | `--track-changes=all` 的独立读取视图 | 可辅助交叉核对，不作为无损往返编辑器。 |

仅抽取普通段落文本、转换 Markdown 或拼接 XML 文本节点，都不足以可靠理解多作者 redline。此前的合成样例已展示：`python-docx` 的普通段落文本可能漏掉修订中的文字；直接拼接文本节点可能漏掉删除内容，或把删除和新增并排拼成错误正文。重写段落还可能使批注锚点与批注对象脱离。

`superdoc-redlines` 的多 Agent 编辑计划也可参考，但其对既有 tracked changes 的读取方式需要单独核对，不作为本项目完整审计的可信读取层。参考：<https://github.com/yuch85/superdoc-redlines>。

## 9. 云环境配置与验证边界

以当前 `package.json`、lockfile 和测试说明为准。分析基线要求 Node `>=22.12.0`，使用 ESM。首批依赖按 lockfile 安装，通常为 `npm ci`；不要运行广泛升级、`npm audit fix --force` 或替换 SDK 来解决无关问题。

应执行的仓库检查包括：

```bash
npm test
npm run format:check
npm run check:release
```

先记录必要的基线，再对修改运行有针对性的回归和仓库要求的最终检查。若安装受网络限制、测试被 runner 阻止或环境缺少组件，记录准确命令、错误和 `ENV_BLOCKED`。缺依赖造成的 skip 不能描述为完整测试通过。

按仓库要求更新 `CHANGELOG.md` 和 `release/TEST-SUMMARY.md`。`release/approved-files.json` 的生成必须遵守最终复核流程；发现 manifest 变化时按规则处理，不跳过 release check、不删除保护来宣称通过。

Cursor-hosted 云环境与用户 Mac 上的真实 Word 环境不同。离线 fixtures、mock、Linux Electron 或历史 native 报告，都不能证明新版本已经通过 Mac 原生验收。

| 云端可以验证 | 需要另行原生验收 |
| --- | --- |
| 证据 schema、分页、映射、覆盖和状态转换 | 当前 Mac Word 的 Office.js 返回值与 API 支持 |
| 合成数据、记录样本、纯逻辑回归 | 真实审阅窗格的修订、作者、批注及回复显示 |
| 只读工具分类、过期处理及权限边界 | Word sidebar 加载、交互与运行时工具调用 |
| 项目离线测试、格式、release 检查 | Electron/macOS launcher、注册、TLS 信任及生命周期 |
| 可复现的合成输出示例 | 后续写入／导出文档保存重开后的编号、格式、字段和批注 |

未执行的原生项标记 `NATIVE_NOT_RUN`。不要为本批启动真实 CodeBuddy／Claude 模型、读取私人配置、修改系统信任或连接真实合同。使用临时合成样本。

如果需要配置 Cursor 环境，优先使用其引导设置；只有确有需要时才增加项目级 `.cursor/environment.json`。保存过的个人／团队环境可能优先于仓库配置，需要检查实际生效环境，而不能仅凭文件存在就认为已加载。

## 10. 本地 Mac 原生验收清单模板

在首批交付中补全与具体实现对应的操作步骤和预期输出，供后续本地会话执行：

1. 使用合成协议，创建至少三位作者的插入／删除／替换，以及批注、回复和 resolved 状态。可加入相同短句出现在不同位置的情况。
2. 在 Word 中确认实际修订与批注，再调用新增只读审计入口；核对作者、before/current、来源和缺口。
3. 准备跨越 `office_read_story`／`office_read_formatting` 页界的样本，核对分页结束与覆盖状态；后续接入有分页的修订／批注适配器时，再增加对应样本。
4. 在读取期间修改文档，核对不一致或过期状态会被识别；无法检测的范围应明确展示。
5. 验证只读操作不会改变任何修订接受状态、批注状态或正文。
6. 记录 macOS／Word 版本、项目 SHA、具体样本、实际结果和证据。后续涉及写入／导出时，再加入保存重开和版式检查。

无法在当前执行环境运行这些步骤时，保留待验收状态，不能复制历史 PASS。

## 11. 每批任务的交付格式

最终报告至少包含：

- 实际分支和提交 SHA，以及本批实现的用户可见行为。
- 关键改动文件与设计取舍。
- 已执行命令与结果：`PASS`、`FAIL`、`ENV_BLOCKED`、`NOT_RUN`；不把未执行项计入通过率。
- 原生 Word 状态：`NATIVE_NOT_RUN` 或本次真实运行证据。
- 已知限制与覆盖缺口。
- 可复现的合成样例和输出。
- 本地 Mac 验收步骤。
- 下一批建议；未经本次范围授权，不直接开始 P1B/P1C/P2/P3。

进度文件应足以让新会话接手，但不要把整个对话日志、敏感文件或个人路径写进仓库。

## 12. 开发约束与范围管理

- 首批优先使用已有依赖，避免为暂不实现的 ZIP importer 提前引入工具链。确需新依赖时，说明用途、固定版本、许可证和替代方案，并按仓库授权规则处理。
- 不直接合并 `main`、发布安装包、增加 GitHub Actions、修改全局配置或升级无关依赖。
- 在 Cursor 机器中执行测试；不为本批主动触发 CI。推送前检查已有工作流触发条件，并尊重用户对 CI 的限制。
- 当前默认 `office-only` 是产品运行时边界，不限制 Cursor 作为开发工具正常阅读仓库、运行已授权的离线测试；也不意味着允许产品启动时加载开发者私人 skills/hooks/MCP。
- 遇到实现细节选择，依据现有代码作出可逆决定并记录，不反复停下来询问。遇到确实缺少授权或阻断条件，先完成不受阻的工作，再明确说明阻塞原因。

## 13. 建议追加到 AGENTS.md 的简短入口

以下内容只作为新增段落，不替换现有文件：

```markdown
## Redline review development

For redline, multi-author review, or multi-submission work, read
`docs/REDLINE-REVIEW-PLAN.md` in full and check
`docs/REDLINE-REVIEW-PROGRESS.md` if present before editing.
Implement only the phase authorized in the current task; record evidence,
coverage gaps, and actual test results. Cloud/offline tests do not establish
native macOS Word compatibility. Preserve the existing Office-only policy,
write approvals, live preconditions, serial execution, and RESULT_UNKNOWN.
```

这段入口负责提醒；详细文件仍需被 Agent 实际读取。无需为同一计划同时维护一份重复的 `.cursor/rules` 全文。

## 14. 后续会话的启动方式

首次：将本文件全文与用户的本次开发指令一起发送给 Cursor，让它落库并实现 P1A。仅发本地文件名或 ChatGPT 的下载链接，不能保证 Cursor 能读取内容。

以后：选择包含已提交规划与进度记录的分支，在任务中明确要求读取这两个文件，并指定当前批次。只有旧分支包含文件时，新任务从 `main` 启动并不会自动获得它们。

本地尚未提交／推送的项目修改也不会自动出现在 Cursor 云端。确认需要的代码已在所选远端分支上；不要让 Agent 为追赶历史分析而重置用户现有工作。

## 15. 参考来源

### 项目与 Cursor

- [Office CodeBuddy 仓库](https://github.com/brucecbi/office-codebuddy)
- [仓库 AGENTS.md](https://github.com/brucecbi/office-codebuddy/blob/main/AGENTS.md)
- [仓库 package.json](https://github.com/brucecbi/office-codebuddy/blob/main/package.json)
- [Cursor Cloud Agents](https://cursor.com/docs/cloud-agent)
- [Cursor Cloud 环境设置](https://cursor.com/docs/cloud-agent/setup)
- [Cursor Rules 与 AGENTS.md](https://cursor.com/docs/rules)

### DOCX 与工具行为的核验入口

- [Microsoft：Word tracked changes](https://learn.microsoft.com/en-us/office/dev/add-ins/word/manage-tracked-changes)
- [python-docx issue 1604](https://github.com/python-openxml/python-docx/issues/1604)
- [OfficeCLI revision 文档](https://github.com/iOfficeAI/OfficeCLI/wiki/word-revision)
- [OfficeCLI issue 279](https://github.com/iOfficeAI/OfficeCLI/issues/279)
- [OfficeCLI PR 487](https://github.com/iOfficeAI/OfficeCLI/pull/487)
- [safe-docx 工具参考](https://github.com/UseJunior/safe-docx/blob/main/packages/docx-mcp/docs/tool-reference.generated.md)
- [Docxodus consolidate 源码基线](https://github.com/JSv4/Docxodus/blob/564881ac409493a89fcf64074dd8b61b1ea651e0/Docxodus/DocxDiffConsolidate.cs)
- [Pandoc tracked changes 选项](https://pandoc.org/MANUAL.html#option--track-changes)

以上链接用于检查能力与限制。issues、PR 和工具版本可能变化；实际集成应固定并记录被验证版本。此前研究中的离线样例或旧测试结果不得计为本次开发验收。
