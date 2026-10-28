# 高阶智能文档排版系统

基于 **OOXML 标准**的 Word 文档排版系统：文档解析、格式规范化、样式调整与自动化排版，
支持模板管理、文档转换、历史记录及实时进度展示。

前后端分离，逐步替代原有 Pandoc 转换流程与 Windows COM 依赖，支持 Linux 跨平台部署。

---

## 技术栈

| 层 | 选型 |
|----|------|
| 后端 | Python 3.11+ / FastAPI |
| OOXML 核心 | python-docx + lxml（直接操作 XML） |
| 排版引擎 | LibreOffice headless + UNO（可选，用于目录页码刷新 / PDF 导出） |
| 前端 | Vite + React 18 + TS + Ant Design 5 + morphicons |

> 设计目标：**无 Office、无 COM、无 Pandoc**，`pip install` 即可在 Linux 运行。

---

## 目录结构

```
doc-format-sys/
├── backend/                    # 后端（FastAPI）
│   ├── app/
│   │   ├── api/v1/             # 路由层（薄，只做校验与编排）
│   │   ├── services/           # 业务服务（一个文件一个类，单一实体）
│   │   ├── domain/ooxml/       # OOXML 核心能力（纯函数式）
│   │   ├── converters/         # 排版引擎抽象（LibreOffice / Noop）
│   │   ├── schemas/            # Pydantic 契约
│   │   ├── core/               # 配置、常量、异常、日志
│   │   └── workers/            # 任务执行器
│   ├── config/                 # 常量外置（JSON，禁止硬编码）
│   ├── resources/              # 内置种子资源（封面模板）
│   ├── tests/                  # 测试与回归
│   └── requirements.txt
├── frontend/                   # 前端（Vite + React + Ant Design）
│   ├── src/constants/          # 文案、文档类型、接口路径
│   └── constants.ts            # 开发端口
├── docs/                       # 文档
│   ├── REFACTOR_PLAN.md        # 重构架构与迁移方案
│   └── REFACTOR_CHANGELOG.md   # 重构变更记录
└── test/                       # 回归样例（src.docx → obj.docx）
```

---

## 运行

先起后端，再起前端。前端开发服务器把 `/api` 代理到 `http://127.0.0.1:8010`。

### 后端

```bash
conda activate doc-format-sys
pip install -r backend/requirements.txt
cd backend
python -m uvicorn app.main:app --host 127.0.0.1 --port 8010
```

- 接口文档：<http://127.0.0.1:8010/docs>
- 健康检查：<http://127.0.0.1:8010/api/v1/health>

### 前端

```bash
cd frontend
npm install
npm run dev
```

浏览器打开 <http://127.0.0.1:5173> 。

页面保留原工作台能力：上传 `.doc` / `.docx`、选择文档类型（初步设计 / 可行性研究及代可研、代初设）、实时进度、下载结果、按同一文件查看转换历史。

### 4. 运行测试

```bash
cd backend
pytest -q
```

---

## 排版引擎

系统把「需要外部排版引擎才能完成」的操作（目录页码刷新、`.doc` 转换、PDF 导出）
隔离在 `DocConverter` 接口后面。

| 引擎 | 说明 |
|------|------|
| `LibreOfficeConverter` | Linux 默认实现，走 `soffice` + UNO |
| `NoopConverter` | 无引擎降级（纯 OOXML 场景、测试） |

通过环境变量 `DFS_LAYOUT_ENGINE` 控制：`auto`（默认，自动降级）/ `libreoffice`（强制，缺失即启动失败）/ `none`。

> **注意**：未安装 LibreOffice 时，目录会以 OOXML 域代码形式插入，但页码不会刷新，
> 需在 Word/WPS 中按 `F9` 更新，或安装 LibreOffice 后由引擎自动刷新。

---

## 配置说明

所有可调参数外置在 `backend/config/*.json`，代码中禁止硬编码：

| 文件 | 内容 |
|------|------|
| `cover_config.json` | 封面模板名、文档类型文字、字段占位、元数据规则 |
| `layout_rules.json` | 表格字体、目录标题关键词、标题层级 |
| `document_settings.json` | 纸张尺寸、页边距、单位换算 |
| `picture_settings.json` | 图片对齐、行距、环绕方式 |

运行期配置（目录、引擎、并发）通过环境变量 `DFS_*` 前缀覆盖，见 `app/core/config.py`。

---

## 文档

- [重构架构与迁移方案](docs/REFACTOR_PLAN.md)
- [重构变更记录](docs/REFACTOR_CHANGELOG.md)

---

## 许可证

[MIT License](LICENSE)
