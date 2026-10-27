# CHANGELOG

> 本文件记录项目的**已改动点**与**待修改内容**。
> 详细设计见 [docs/REFACTOR_PLAN.md](docs/REFACTOR_PLAN.md)，
> 重构明细见 [docs/REFACTOR_CHANGELOG.md](docs/REFACTOR_CHANGELOG.md)，
> 排障过程见 [docs/TROUBLESHOOTING-20261027-232006.md](docs/TROUBLESHOOTING-20261027-232006.md)。
>
> 最近更新：2026-10-27 23:20

---

## 一、已改动的点

### 1. 删除重构前的旧代码

| 删除对象 | 说明 |
|----------|------|
| `utils/`（46 个文件） | pandoc 转换、COM 操作、各 step 散落脚本、重复的 `converter.py`、时间戳备份文件 |
| `service/`（4 个文件） | 5400 行 `converter.py` 上帝类、`dify_service.py` 远端依赖 |
| `sqls/`（2 个文件） | `doc_node.sql`、`doc_rec.sql`，新实现已不使用数据库 |

> 均已纳入 git 跟踪，可从历史恢复。

### 2. 后端重构为分层结构（`backend/`）

```
app/api/v1/          路由层（薄，只做校验与编排）
app/services/        业务服务（一个文件一个类，单一实体）
app/domain/ooxml/    OOXML 核心能力（纯函数式，不感知 HTTP/DB）
app/converters/      排版引擎抽象（LibreOffice / Noop）
app/core/            配置、常量、异常、日志
app/workers/         任务执行器
config/              常量外置（JSON，禁止硬编码）
resources/           内置种子资源（封面模板）
```

- 20 步可枚举流水线：`open → repair → numbering → orientation → footer → highlight → template → toc_cleanup → headings → split → cover → content → merge → toc_insert → toc_section → picture → finalize → save → engine → validate`
- 去 Pandoc、去 Windows COM

### 3. 封面模板整合（24+7 → 1+1）

**改动前**：按 `文档类型 × 日期型 × 长度档` 维护 24 份模板 + 7 份配置。

**改动后**：

```
backend/resources/cover_template.docx   唯一封面模板
backend/config/cover_config.json        唯一封面配置
template/archive/                       24 份旧模板 + 7 份旧配置归档
```

逐份 diff 后确认 24 份模板只差三处，全部改为运行时回填：

| 维度 | 差异 | 处理 |
|------|------|------|
| 文档类型 1/2/3/4 | 文本（初步设计 / 可行性研究 / （代可研）/（代初设）） | 运行时替换，副行在类型 1/2 时整段删除 |
| 日期型 | 年月日 / 年月 | 从源封面自动识别 |
| 长度档 1/2/3 行 | 只差 1 个空段落 | 项目名自动折行，无需区分 |

配套改动：

- `app/domain/ooxml/cover.py` 重写为单一配置驱动
- `app/services/cover_service.py` 重写，删除模板选择逻辑
- `app/core/constants.py` 新增 `cover_config()`；`cover` 配置从 `layout_rules.json` 迁出
- `app/domain/ooxml/tables.py` 改用 `cover_config()`
- 新增 `use_cover_template` 选项：合成文档测试时走就地清理，不强制替换模板
- `app/core/config.py` 新增 `resources_dir`

### 4. 修复两个真实 Bug

**① 人员识别空格 Bug**

源稿写的是「总 设 计 师：曹 永」，职务带空格导致关键词匹配失败，封面第三位人员一直留空。
修复：职务/姓名先去空格再匹配。现在封面三人（王刚 / 孙海权 / 曹永）全部正确回填。

**② 悬空样式 Bug**

封面包与正文包样式表不同，合并后正文引用的 `a` / `a0` / `a1` / `Normal` 等样式悬空，
`validate` 步骤长期失败。
修复：`DocumentMerger._merge_styles()` 合并源样式到目标 `styles.xml`。现在 `dangling_styles = 0`。

### 5. 接入 LibreOffice 排版引擎

**安装**：LibreOffice 26.8.1.1（winget）。

**修复 `bootstrap.ini` 报「已损坏」**：

- 真因是文件被写成 **UTF-8 BOM**（`EF BB BF` 开头），LO 要求首字节为 `[`
- 同时 `InstallMode=<installmode>` 占位符未被 winget 替换
- 修复：管理员权限 → 纯 ASCII 无 BOM + CRLF 写入，`InstallMode=install`

**重写域刷新链路**（`app/converters/libreoffice.py`）：

关键结论（实测）：**由 LibreOffice 自带 Python 启动 soffice 会失败**（进程起来但端口不监听），
必须由主进程启动、LO Python 只负责连接。

```
后端 Python：启动 soffice --accept=...ServiceManager → 等端口 → 调脚本 → taskkill /T 清理
LO Python ：连接端口 → 刷新索引 → store() → 关闭
```

其他要点：

- **监听端**用 `StarOffice.ServiceManager`，**连接端**用 `StarOffice.ComponentContext`，不可混用
- `uno` 模块只在 LibreOffice 自带 Python 中存在，conda Python 无法 `import uno`
- 新增 `scripts/refresh_fields.py`（只连接模式，支持 `--port` / `--timeout`）
- 新增 `_resolve_python` / `_free_port` / `_wait_for_port` / `_terminate_tree` / `_find_in_common_locations`

**结果**：目录页码可自动刷新，生成 **140 条**带页码条目（与 obj.docx 对齐）。

### 6. 修复流水线顺序 Bug

`save` 原本排在 `engine` 之后，会把引擎就地刷新后的文件用内存包重新覆盖，
导致目录刷新结果丢失。已将 `save` 移到 `engine` 之前。

### 7. 文档

| 文件 | 说明 |
|------|------|
| `README.md` | 重写（原为 Pandoc/COM 版本，已过时） |
| `docs/REFACTOR_CHANGELOG.md` | 重构变更记录 |
| `docs/TROUBLESHOOTING-20261027-232006.md` | 目录页码刷新完整排障记录 |
| `template/说明.txt` | 更新为封面整合后的说明 |

### 8. 测试

- 33 项测试全部通过（含新增回归测试）
- 新增 `tests/test_regression.py`：`test/src.docx → obj.docx` 结构级对比
- `tests/test_format_pipeline.py`：改用 `PIPELINE_STEPS` 断言步骤数，默认 `use_cover_template=False`

---

## 二、待修改的内容

### P0 — 阻断交付，必须修复

#### 1. 接入引擎后图片丢失 ⚠️ 新引入的回归

| 项 | 当前输出 | obj.docx |
|----|----------|----------|
| `word/media/` 部件 | **0** | 2 |
| 段落数 | 779 | 914 |

**现象**：LibreOffice `document.store()` 重新打包 docx 后，图片部件全部丢失。

**可能原因**（未验证）：
- 图片为 WMF 等旧格式，LO 保存时转换失败
- 关系（`document.xml.rels`）或 `[Content_Types].xml` 不完整
- 包内部件顺序/兼容性问题

**建议方案**：
1. 先定位 LO store 丢图的具体原因
2. 若无法保留，改为**只回填刷新后的 TOC 段落**到原文档，不用 LO 整体保存，
   避免动到图片与其它部件

**注意**：在此问题修复前，`engine` 步骤的收益（目录页码）不足以抵消图片丢失的损失，
建议临时设 `DFS_LAYOUT_ENGINE=none` 保持图片完整，或按方案 2 改造。

#### 2. 回归对比指标口径

`diff.py` 的「正文文本保留度」把目录也算进正文比对：

- obj 的目录是 140 条带页码条目，输出是 1 条域占位时，相似度被拉低到 85.9%
- 剔除目录区后实测 **正文相似度 100%**

**建议**：改为「剔除目录区后再算正文相似度」，让指标真实反映正文有无丢失。

### P1 — 功能待确认

#### 3. 封面「档案号 / 图号」业务规则

封面上的 `10795-436245`、`36830-446405` 目前是模板里的**静态值**（旧系统每份模板也各自写死）。

**待确认**：这两个编号是按文档生成/分配的，还是固定值？若需动态生成，请提供规则。

#### 4. 目录条目数差异

| 项 | 输出 | obj.docx |
|----|------|----------|
| TOC 条目 | 140 | 127 |

**待确认**：是标题识别多标了，还是 obj 的目录本身被人工删减过。

#### 5. 段落数差异

输出 779 段 vs obj 914 段（差 135）。

**推测**：obj 是经 Word 刷新过域的版本，TOC 条目以不同结构存储；需在图片问题修复后重新比对。

### P2 — 暂不处理（已确认）

#### 6. 横向节数 1 vs 2

原 COM 流程在横/竖切分时会额外生成一个横向节。
**业务方已确认**：这是原 COM 实现的 bug，无需复现。

### P3 — 待建设

#### 7. 前端

Vite + React 18 + TS + Ant Design 5 尚未搭建（`docs/REFACTOR_PLAN.md` §3.2 已定选型）。

#### 8. 数据库

当前**不使用数据库**：

- 历史记录 → `HistoryService`，JSONL 追加文件 `backend/data/history/records.jsonl`
- 任务状态 → `TaskService`，内存 + JSONL 落盘

留痕功能已可用。若后续需要多实例/复杂查询再引入 PostgreSQL（需单独评审）。

#### 9. `pytest` 弃用告警

`starlette.testclient` 提示 `httpx` 用法将弃用，建议迁移到 `httpx2`。非阻塞。

---

## 三、当前状态速览

| 项 | 状态 |
|----|------|
| 测试套件 | ✅ 33 项通过 |
| 封面整合（24+7 → 1+1） | ✅ 完成，内容与 obj 对齐 |
| 目录页码刷新 | ⚠️ 可刷新（140 条），但会丢图片 |
| 图片保留 | ❌ 接入引擎后丢失（P0） |
| 前端 | ⬜ 未开始 |
| 常量外置 | ✅ 全部迁入 `backend/config/*.json` |

---

## 四、环境依赖备忘

| 依赖 | 版本 / 路径 | 说明 |
|------|-------------|------|
| Python | 3.11（conda 环境 `doc-format-sys`） | 主运行时 |
| LibreOffice | 26.8.1.1，`C:\Program Files\LibreOffice` | 需手动修复过 `bootstrap.ini` |
| LO Python | `C:\Program Files\LibreOffice\program\python.exe` | 唯一能 `import uno` 的解释器 |

> **部署到 Linux 时**：安装 LibreOffice 官方包并额外安装 `python3-uno`，
> 引擎会自动探测 `/usr/bin/soffice` 与 `/usr/bin/python3`，无需改代码。
> 另需随镜像安装中文字体并配置 `fontconfig` 别名，否则分页会错位
> （见 `docs/REFACTOR_PLAN.md` §6.3）。

---

## 五、开源许可

当前全部依赖均为开源许可，无商业授权风险：

| 组件 | 许可 |
|------|------|
| FastAPI / Uvicorn | MIT / BSD |
| python-docx | MIT |
| lxml | BSD |
| Pydantic | MIT |
| Pillow | MIT-CMU |
| LibreOffice | **MPL 2.0**（弱 copyleft，允许商业闭源调用，仅要求修改其源码时开源） |
| pytest / httpx | MIT / BSD |

> 唯一需注意的是 LibreOffice 的 MPL 2.0：只要**不修改 LibreOffice 源码**（仅以独立进程调用），
> 商业使用没有额外义务。将来若换成 Aspose.Words，则为商业授权，需单独评估。
