# 高阶智能文档排版系统 — 重构架构与迁移方案

> 状态：草案（Date 2026-10-27）
> 范围：技术选型结论、目标架构、转换引擎替换、迁移路径
> 约定：本文档只做设计决策记录，不包含代码改动

---

## 1. 项目定位

**高阶智能文档排版系统**，采用前后端分离架构，基于 OOXML 标准实现 Word 文档解析、格式规范化、样式调整与自动化排版，支持模板管理、文档转换、历史记录及实时进度展示。支持 Linux 跨平台部署，逐步替代原有 Pandoc 转换流程与 Windows COM 依赖。

### 1.1 三个必须达成的目标

| 编号 | 目标 | 验收标准 |
|------|------|----------|
| G1 | 去 Windows 依赖 | 生产环境无 `pywin32`、无 Office、无 `.exe` 二进制；`pip install` 可在 Linux 完成 |
| G2 | 去 Pandoc 转换链路 | 文档转换不再经过 pandoc，转换后无需"补格式"步骤 |
| G3 | 排版结果可校验 | 每次输出通过结构断言 + golden-file diff，失败必须显式报错，不允许静默交付 |

### 1.2 非目标

- 不做在线协同编辑（不走 OnlyOffice/Collabora 的编辑器形态）
- 不做 Word 100% 保真渲染（明确边界见 §6.4）
- 不在本期做多租户隔离

---

## 2. 现状诊断

重构的前提是明确"旧结构为什么必须废弃"。以下均为代码事实，不是主观判断。

### 2.1 COM 依赖的不可控性

`service/converter.py` 中存在 39 处以上 COM 调用（`DispatchEx` / `Documents.Open` / `Fields.Update` / `Quit`），并伴随两个"对抗性"补丁：

```91:91:service/converter.py
def _kill_existing_winword_processes():
```

```108:108:service/converter.py
def _wait_for_file_release_for_step(path: str, timeout: float = 10.0, interval: float = 0.3) -> bool:
```

需要"杀掉 Word 进程"和"轮询等待文件锁释放"才能推进流程，说明架构已经在与 Word 的进程模型对抗。根因是模型决定的，不是 bug：

- Word 是单实例 COM Server，文档级独占锁
- RPC 调用超时、`gencache` 缓存残留、进程泄漏均不可控
- 只能运行在 Windows，本地与生产行为无法一致

**结论：COM 不是"需要优化"，是"必须移除"。**

### 2.2 Pandoc 链路的方向性错误

`service/converter.py:2851` 的 `step2_pandoc_convert` 之后，仍有大量格式修复步骤。原因是 pandoc 是**语义转换器**，设计上刻意丢弃排版信息（分节属性、页眉页脚引用、编号细节、文本框、图片锚定方式）。

它在本流程中的角色是"制造格式问题，再由后续 step 修复"。移除它可以直接删掉一批补丁式步骤。

### 2.3 全量 COM 依赖分布

| 文件 | COM 相关命中数 |
|------|----------------|
| `service/converter.py` | 95 |
| `utils/converter.py` | 57 |
| `utils/pandoc_converter.py` | 44 |
| `utils/document_orientation_processor.py` | 15 |
| `utils/docx_supplement.py` | 12 |

`service/` 与 `utils/` 存在**功能重复**（两份 `converter.py`），重构时必须合并为单一实体（符合"service 单类单实体"约定）。

### 2.4 其他待清理项

- `utils/pandoc.exe`：二进制入库，应改为镜像构建产物，不进 git
- `requirements.txt`：`Flask` 与 `FastAPI` 并存，统一到 FastAPI
- `utils/picture_format_keeper.py` 使用 `PIL.ImageGrab`（Windows 截图），需替换为引擎渲染
- `utils/document_orientation_processor - 2026年5月25日082935.py`：带时间戳的备份文件，直接删除

---

## 3. 技术选型决策

### 3.1 后端语言：Python + FastAPI

**决策：Python 3.11+ / FastAPI。**

先要区分项目的两段工作，语言只影响 A 段：

- **A. 修改 OOXML 结构** —— 改样式、页边距、页眉页脚、表格、编号、图注。这是"改 XML"，用库完成。
- **B. 计算排版结果** —— 分页落在哪、目录页码是多少、域的结果是什么、导出 PDF 预览。**这不是任何库能算出来的**，必须有真正的排版引擎。

B 段无论选 Python、Java 还是 C#，最终都是驱动 LibreOffice。这是同一份成本，不参与语言比较。

#### 候选对比

| 语言 / 栈 | 改 OOXML 的能力 | Linux 可用 | 与现有资产关系 | 结论 |
|-----------|----------------|-----------|---------------|------|
| **Python + python-docx + lxml** | 高层 API 覆盖不全，但**可随时下钻 lxml 直改 XML**，过渡无缝 | 是 | `utils/` 上百个脚本可直接搬运 | **选用** |
| C# / .NET 8 + Open XML SDK | 上限最高：官方强类型 schema、Open-Xml-PowerTools 覆盖图片替换/表头重复/搜索替换；配合 Aspose.Words 可省掉 LibreOffice | 是（纯托管） | 需全部重写 + 换技术栈 | 次选，见 §3.3 |
| Java + Apache POI(XWPF) | 可用，但 TOC 页码、多节不同页眉、奇偶页、文本框 `w:txbxContent`、编号重排需大量下钻 XML —— 恰是核心需求 | 是 | 需全部重写 | 可行但无优势 |
| Node/TS | `docx` 库强项是从零生成；改现有文档需自建 OPC 包、关系、样式继承、编号定义 | 是 | 需重写，且要自造底层 | 不选 |
| Go | `unioffice` 商业收费，开源 `gooxml` 已停维护 | 是 | — | 不可行 |
| Rust | 仅能生成简单文档 | 是 | — | 不可行 |
| C++ | 无成熟 OOXML schema 库，等于自研 LibreOffice 子集 | 是 | — | 不可行 |

#### 为什么不选语言性能

项目瓶颈是 LibreOffice 进程、磁盘 IO、大文件内存，不是解释器速度。一个 50 页 docx 的解析在 Python 是数百毫秒级，C++ 省下的是几十毫秒，代价是数月工期。**这里谈 C++ 是纯负收益。**

#### 关于 Next.js 的单独说明

Next.js 不适合承担本项目后端，原因是产品形态而非性能：

1. RSC / Server Actions 假设前后端一体，与"前后端分离 + 各自独立镜像"直接冲突
2. API Route 跑在单进程 Node，CPU 密集的转换任务会阻塞事件循环；多副本时又需外置任务状态，等于自建 worker，Next 只剩路由壳
3. 真正的重活交给外部进程（LibreOffice），与 Web 框架无关
4. 改文档需要自建 OPC 层，生态是空档

**结论：前端用 Vite + React，不引入 Next.js。**

### 3.2 前端：Vite + React + TS + Ant Design

与 `ASE-Weekly-Scheduling` 保持一致：

- React 18 + TypeScript + Vite 5
- Ant Design 5（遵循 Ant Design 视觉语言）
- 状态管理 Zustand，请求 axios
- 图标统一走 morphicons 库
- 路由 React Router 6

**不选 Next.js** 的理由见 §3.1。

### 3.3 留一个逃生出口：Aspose.Words

如果后期发现 LibreOffice 保真度或性能无法满足，可切换到 `Aspose.Words`（.NET / Java / Python 三版均有）：

- 纯库调用，**无外部进程、无 profile 锁冲突**
- 分页计算、PDF 导出、TOC 更新一步到位
- 代价：商业授权

架构上通过 `DocConverter` 接口隔离，届时只替换一个实现类，业务代码零改动（见 §4.3）。

### 3.4 选型汇总

| 层 | 选型 |
|----|------|
| 前端 | Vite + React 18 + TS + Ant Design 5 |
| 后端 | Python 3.11+ / FastAPI |
| OOXML 核心 | python-docx + lxml（直接操作 XML） |
| 排版引擎 | LibreOffice headless + UNO |
| 任务队列 | Redis + worker（见 §5） |
| 数据库 | PostgreSQL（或沿用 MySQL），SQLAlchemy 2 + Alembic |
| 部署 | Docker + Nginx，Linux |

---

## 4. 目标架构

### 4.1 目录结构

对齐 `ASE-Weekly-Scheduling` 的组织方式：

```text
doc-format-sys/
├── fe/                            前端
│   ├── src/
│   │   ├── api/                   OpenAPI 生成的客户端
│   │   ├── components/
│   │   ├── pages/
│   │   ├── stores/
│   │   ├── constants/             前端常量（禁止硬编码）
│   │   └── types/
│   └── package.json
├── doc-format-be/                 后端
│   ├── app/
│   │   ├── api/v1/                路由层（薄，只做参数校验与编排）
│   │   ├── services/              业务服务（一个文件一个类，单一实体）
│   │   ├── domain/ooxml/          OOXML 核心能力
│   │   │   ├── styles.py          样式与字体规范化
│   │   │   ├── sections.py        分节、页面设置、横竖向
│   │   │   ├── header_footer.py   页眉页脚
│   │   │   ├── tables.py          表格规范与拆分
│   │   │   ├── numbering.py       编号与图注
│   │   │   ├── toc.py             目录域与标题
│   │   │   └── pictures.py        图片尺寸与锚定
│   │   ├── converters/            转换引擎抽象（见 §4.3）
│   │   ├── templates/             模板管理领域
│   │   ├── workers/               队列任务与进度事件
│   │   ├── models/                SQLAlchemy 模型
│   │   ├── schemas/               Pydantic 契约
│   │   └── core/                  配置、日志、异常、常量
│   ├── config/                    常量与参数外置（JSON / TOML）
│   ├── alembic/
│   └── pyproject.toml
├── deploy/                        docker / nginx / systemd
├── docs/                          本文档所在
├── scripts/                       构建与发布脚本
├── tests/
│   ├── golden/                    golden-file 基线
│   └── fixtures/                  输入样本文档
└── artifacts/                     gitignore
```

### 4.2 分层原则

1. **路由层不写业务**：`api/v1/` 只做鉴权、参数校验、调用 service、组装响应
2. **service 单类单实体**：禁止 `converter.py` 这种 5000 行、39 个 step 的上帝类（现 `service/converter.py` 约 5400 行）
3. **OOXML 能力下沉到 `domain/ooxml/`**：纯函数式，输入 `Document` 对象/XML 节点，不感知 HTTP、不感知数据库
4. **常量外置**：所有格式常量（字号、行距、缩进、页边距默认值、编号规则）写入 `config/`，代码内不得出现魔法值
5. **复用优先**：`domain/ooxml/` 的能力必须可被 service 和 worker 共同调用，不允许复制粘贴

### 4.3 转换引擎抽象

```text
DocConverter (Protocol / ABC)
├── LibreOfficeConverter     Linux 默认实现，走 UNO
├── LegacyComConverter       Windows 回退，逐步下线
└── (AsposeConverter)        备选，纯库无进程
```

接口建议最小集：

```python
class DocConverter(Protocol):
    def probe(self) -> ConverterCapabilities: ...            # 支持哪些能力
    def convert(self, src: Path, dst: Path, fmt: str) -> None: ...
    def refresh_fields(self, path: Path) -> None: ...        # TOC / 页码 / 交叉引用
    def page_count(self, path: Path) -> int: ...
    def export_pdf(self, src: Path, dst: Path) -> None: ...
    def render_page(self, src: Path, page: int, dst: Path) -> None:  # 预览图
```

设计要点：

- 业务代码只依赖接口，不依赖具体引擎 → 迁移可度量、可回滚
- `probe()` 用于启动自检：引擎缺失或版本不符时**启动即失败**，不等到用户上传才报错
- 保留 `LegacyComConverter` 的意义是**灰度替换**，不是长期共存；规定下线时点

---

## 5. 长任务与进度

文档转换是分钟级、吃满 CPU 的任务，必须与 Web 进程隔离。

### 5.1 模型

```text
FE ──POST /api/v1/tasks──▶ API ──enqueue──▶ Redis ──▶ Worker(独立容器)
 │                                                    │
 └──GET /api/v1/tasks/{id}/events (SSE)◀──进度事件─────┘
```

- **API 进程绝不执行转换**，只做入队与查询
- Worker 独立容器/进程，可水平扩容；并发上限由转换引擎决定（见 §6.2）
- 任务状态落库（`doc_task` / `doc_task_step` / `doc_task_artifact`），Redis 只做队列与事件总线

### 5.2 为什么用 SSE 而不是 WebSocket

文档任务的进度是**单向流**（服务端 → 客户端），SSE 走普通 HTTP，天然兼容 Nginx 反代、无需心跳与重连协议、无状态更易横向扩容。WebSocket 在这里增加一半运维负担而无收益。

### 5.3 进度粒度

进度事件按**步骤**上报，不按百分比臆造：

```json
{"task_id": "...", "step": "format_tables", "index": 4, "total": 11, "status": "running"}
```

每一步必须可独立重试，且步骤名与 `domain/ooxml/` 能力一一对应，便于定位问题。

---

## 6. LibreOffice 作为排版引擎

### 6.1 能力映射（COM → UNO）

| 当前 COM 调用 | UNO 等价 |
|---------------|----------|
| `win32.DispatchEx("Word.Application")` | 连接常驻 `soffice` 的 UNO socket |
| `word.Documents.Open(...)` | `desktop.loadComponentFromURL(...)` |
| `doc.Fields.Update()` | `document.refresh()` / 目录 `update()` |
| `ComputeStatistics(wdStatisticPages)` | `getCurrentController().getPageCount()` |
| `ExportAsFixedFormat` (PDF) | `storeToURL(..., FilterName="writer_pdf_Export")` |
| `PIL.ImageGrab` 截图 | `storeToURL` 导出 PNG / 渲染预览图 |
| `doc.SaveAs2` | `storeToURL` / `storeAsURL` |

`.doc → .docx`、`.wmf` 等旧格式兼容也由引擎处理，无需额外组件。

### 6.2 并发：必须做 profile 隔离

LibreOffice 默认共享 user profile，多进程并发会直接报 **"LibreOffice is already running"** 并退出。这不是"偶发冲突"，是默认行为。

**必须**为每个任务分配独立 profile：

```text
-env:UserInstallation=file:///tmp/lo-profile-<task_uuid>
```

任务结束清理目录。或采用常驻 UNO 池 + 并发限流。业界标准做法（unoconv / JODConverter）均如此。

**架构约束：转换引擎必须独立为 worker 池，并发数显式配置，禁止在 web 进程内调用。**

### 6.3 字体：头号事故源

Linux 无宋体 / 黑体 / 仿宋_GB2312 / 楷体，LibreOffice 会**静默替换**字体 → 行高变化 → 分页错位 → 目录页码全错。

必须做到：

1. 字体文件随镜像构建安装（来源需授权合规）
2. 配置 `fontconfig` 别名映射（如 `宋体 → SimSun`、`仿宋_GB2312 → FangSong`）
3. 输出后校验：比对文档声明的字体名与实际可用字体，发现回退即告警
4. 在 golden-file 测试中固化字体环境版本号

### 6.4 保真度边界（必须对业务方明示）

以下场景转换结果可能与 Word 存在差异，需建立明确的处理策略：

- SmartArt / 图表
- 部分 OLE 嵌入对象
- 罕见编号定义（多层 `numbering.xml`）
- `.wmf` 等旧图元文件（现仓库 `template/tem/media/image1.wmf` 即属此类）
- 域代码的复杂嵌套

**策略：不静默交付。** 每次输出过结构断言（§7），不通过则任务标记 `needs_review`，人工介入。

---

## 7. 质量保障：先建基线，再重构

**这是整个重构最重要的前置条件。** 没有回归基线，任何重构都是盲改。

### 7.1 golden-file 基线

利用现有资产建立基线（`template/` 下已有 `reference_content.docx` 与各类 `reference_cover*.docx`，`cover_replace_config*.json` 可作输入参数）：

1. 用**当前实现**跑一组固定输入，产出并固化结果
2. 结果按 OOXML 归一化后存档（剥离时间戳、`docProps`、随机 ID）
3. 重构后每个 service 抽出来都对着基线跑 diff

### 7.2 差异比较策略

| 层级 | 比较对象 | 判定 |
|------|----------|------|
| XML 结构 | `document.xml` / `styles.xml` / `numbering.xml` 归一化后 | 精确匹配 |
| 语义 | 节数、样式名集合、页眉页脚引用、图片数、表格数、段落数 | 精确匹配 |
| 分页 | 引擎计算的页数与分页位置 | 精确匹配 |
| 视觉 | 导出 PDF 逐页转图后像素比对 | 阈值内允许差异 |

### 7.3 结构断言清单（每个输出必须通过）

- 节数量与页面方向序列符合预期
- 样式名集合未引入未定义样式
- 页眉/页脚关系（`header*.xml` / `footer*.xml`）引用完整，无悬空 `rId`
- 图片关系数与 `media/` 文件数一致
- 编号定义无断裂引用
- TOC 域已刷新，页码非 `0`/占位
- 文档未被标记为"需要修复"（可用 libreoffice 重新打开验证）

---

## 8. 数据模型（草案）

不修改数据库前需单独确认。以下为结构建议：

| 表 | 用途 | 关键字段 |
|----|------|----------|
| `template` | 模板管理 | `id`、`name`、`type`(cover/content)、`variant`、`file_path`、`config_json`、`version` |
| `template_config` | 模板替换规则（对应现有 `cover_replace_config*.json`） | `template_id`、`kind`、`rules_json` |
| `doc_task` | 转换任务 | `id`、`template_id`、`status`、`progress`、`error_code`、`created_at` |
| `doc_task_step` | 步骤明细 | `task_id`、`step`、`status`、`started_at`、`finished_at`、`message` |
| `doc_task_artifact` | 输入输出产物 | `task_id`、`role`(input/output/preview)、`path`、`size`、`sha256` |
| `doc_task_event` | 进度事件（SSE 兜底回放） | `task_id`、`seq`、`payload_json` |

现状 `sqls/doc_rec.sql`、`sqls/doc_node.sql` 可作为历史记录表的参考，但需按新模型重写。

---

## 9. 迁移路径

分五阶段，**每阶段可独立验收、可回滚**。

### P0 — 冻结与基线（工期基准）

1. 现有代码整体提交一次（重构前快照）
2. 清理入库二进制与备份文件（`utils/pandoc.exe`、`* - 2026年5月25日082935.py`）
3. 建立 `tests/fixtures/` 输入样本集与 `tests/golden/` 基线
4. 定下归一化比较工具

**验收：基线可一键复现，diff 工具可运行。**

### P1 — 骨架与新工程

1. 建立 `fe/` 与 `doc-format-be/` 目录骨架
2. FastAPI 基础：配置、日志、异常、路由分层、OpenAPI
3. 前端基础：路由、布局、API 客户端生成、常量文件
4. CI：lint + test + build

**验收：空实现可跑通"上传 → 入队 → 进度 → 完成"的端到端假流程。**

### P2 — OOXML 核心下沉（纯库，不依赖引擎）

把现有 `utils/` 逻辑按能力拆入 `domain/ooxml/`，每个能力配 golden 测试：

1. 样式与字体规范化
2. 分节与页面设置（含横竖向拆分 —— 对应现有 `document_orientation_processor`）
3. 页眉页脚
4. 表格规范与拆分
5. 编号与图注
6. 目录与标题
7. 图片尺寸与锚定

**验收：每个能力有单测 + golden 断言，全部通过。**

### P3 — 引擎替换（核心阶段）

1. 实现 `LibreOfficeConverter`（UNO + profile 隔离 + 字体配置）
2. 建立 worker 池与并发限流
3. 对齐 §6.1 能力映射，逐条替换 COM 调用
4. `LegacyComConverter` 保留为回退开关，仅在 Windows 开发机可选

**验收：Linux 上完整跑通全流程，输出通过 §7.3 全部断言。**

### P4 — 下线旧链路

1. 移除 pandoc 相关步骤与 `pandoc.exe`
2. 移除 `pywin32`、`PIL.ImageGrab` 等 Windows 依赖，从 `requirements.txt` 彻底删除
3. 删除 `LegacyComConverter`
4. 清理旧 `service/converter.py` 与重复的 `utils/converter.py`

**验收：`grep -r "win32com\|pandoc"` 无业务代码命中；`pip install` 在纯净 Linux 镜像完成。**

---

## 10. 风险清单

| 风险 | 影响 | 应对 |
|------|------|------|
| 字体缺失导致分页错位 | 目录页码全错，交付失败 | §6.3 字体随镜像 + 断言校验 |
| LibreOffice profile 冲突 | 并发任务随机失败 | §6.2 每任务独立 profile + 并发限流 |
| 保真度差异（SmartArt/OLE/wmf） | 输出与 Word 不一致 | §6.4 断言拦截 + `needs_review` 人工介入 |
| 无回归基线 | 重构等于盲改 | P0 必须先完成，未完成不得进入 P2 |
| 接口契约漂移 | 前后端联调返工 | OpenAPI 生成 TS 客户端，禁止手写类型 |
| 上帝类回潮 | 重蹈 `service/converter.py` 覆辙 | code review 卡口：单文件超阈值即拒绝 |
| 旧逻辑遗漏 | 功能回退 | P2 阶段按能力逐项对照旧 step 清单 |

---

## 11. 待确认事项

以下需要在实施前与业务方确认，不可臆断：

1. 数据库选型：沿用现有 MySQL，还是 PostgreSQL？（涉及迁移，需单独批准）
2. 字体授权：镜像内可分发的字体清单与授权来源
3. 保真度容忍度：哪些差异业务方可接受、哪些必须人工复核
4. 模板管理的交互形态：版本化模板是否需要回滚与对比
5. 历史记录保留策略：产物文件保留时长与清理规则
6. 并发目标：峰值同时转换任务数（决定 worker 池规模与部署规格）

---

## 12. 决策记录

| 日期 | 决策 | 理由 |
|------|------|------|
| 2026-10-27 | 后端选用 Python + FastAPI | OOXML 可下钻 lxml、现有资产可复用、生态齐全、Linux 原生 |
| 2026-10-27 | 前端选用 Vite + React + AntD | 与 `ASE-Weekly-Scheduling` 一致，无需 Next.js 的同构能力 |
| 2026-10-27 | 不采用 Next.js 作为后端 | 与前后端分离冲突、CPU 密集任务阻塞、改文档需自建 OPC 层 |
| 2026-10-27 | 排版引擎选用 LibreOffice + UNO | Linux 原生、跨平台一致、能真正计算分页与域 |
| 2026-10-27 | 转换引擎抽象为接口 | 支持灰度替换与将来切换到 Aspose.Words |
| 2026-10-27 | 进度采用 SSE | 单向流，运维成本低于 WebSocket |
| 2026-10-27 | 先建 golden-file 基线再重构 | 无基线则重构不可验证 |
