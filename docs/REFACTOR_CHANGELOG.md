# 重构变更记录

> 本文档记录**本次重构实际做了什么**，与 [REFACTOR_PLAN.md](./REFACTOR_PLAN.md)（设计决策/迁移方案）互补。
> 约定：REFACTOR_PLAN 是"计划"，本文档是"已落地的变更"。

---

## 1. 一句话总结

把「Pandoc + Windows COM + 5400 行上帝类 + 44 个散落脚本」的旧链路，重构成
「FastAPI + python-docx/lxml 直改 OOXML + 分层服务 + 20 步可枚举流水线」，
并完成了封面模板的整合收敛与功能回归验证。

---

## 2. 目录结构变化

### 旧结构（已删除）

```
doc-format-sys/
├── main.py                      # 单文件入口（硬编码路径）
├── service/                     # 3 个文件，含 5400 行 converter.py 上帝类
│   ├── converter.py             # 39+ 处 COM 调用、Pandoc 链路
│   ├── convert_history_service.py
│   └── dify_service.py          # 依赖远端 Dify 服务
├── utils/                       # 44 个散落脚本，功能重复
└── sqls/                        # 历史表结构参考（本次保留，待确认）
```

### 新结构

```
backend/
├── app/
│   ├── main.py                  # FastAPI 入口
│   ├── api/v1/                  # 路由层（薄，只做校验与编排）
│   │   ├── health.py
│   │   ├── engines.py
│   │   ├── documents.py
│   │   ├── tasks.py
│   │   ├── templates.py
│   │   └── history.py
│   ├── services/                # 业务服务（一个文件一个类，单一实体）
│   │   ├── format_service.py    # 主流水线编排
│   │   ├── cover_service.py     # 封面构建
│   │   ├── template_service.py  # 模板管理
│   │   ├── upload_service.py    # 上传
│   │   ├── history_service.py   # 历史记录
│   │   └── task_service.py      # 任务
│   ├── domain/ooxml/            # OOXML 核心能力（纯函数式，不感知 HTTP/DB）
│   │   ├── package.py           # 文档包（OPC 读写、关系、部件）
│   │   ├── document.py          # 拆分/合并（含关系重映射、样式合并）
│   │   ├── styles.py            # 样式解析与字体规范化
│   │   ├── sections.py          # 分节、页面方向
│   │   ├── header_footer.py     # 页眉页脚
│   │   ├── tables.py            # 表格规范与封面表格
│   │   ├── numbering.py         # 编号
│   │   ├── toc.py               # 目录域插入与清理
│   │   ├── headings.py          # 标题识别与层级纠正
│   │   ├── highlighting.py      # 高亮清理
│   │   ├── pictures.py          # 图片对齐
│   │   ├── cover.py             # 封面元数据抽取与回填
│   │   ├── repair.py            # 结构修复
│   │   ├── diff.py              # 结构对比（回归工具）
│   │   └── validate.py          # 输出结构断言
│   ├── converters/              # 排版引擎抽象
│   │   ├── base.py              # DocConverter 接口
│   │   ├── libreoffice.py       # LibreOffice 实现
│   │   ├── noop.py              # 无引擎降级实现
│   │   └── registry.py          # 引擎选择
│   ├── schemas/                 # Pydantic 契约
│   ├── core/                    # config / constants / errors / logging / namespaces
│   └── workers/                 # 任务执行器
├── config/                      # 常量外置（禁止硬编码）
│   ├── cover_config.json        # 封面配置（本次新增）
│   ├── layout_rules.json        # 排版规则
│   ├── document_settings.json   # 页面设置
│   └── picture_settings.json    # 图片设置
├── resources/
│   └── cover_template.docx      # 唯一封面模板
├── tests/
│   ├── test_api.py
│   ├── test_format_pipeline.py
│   ├── test_ooxml_modules.py
│   ├── test_package.py
│   └── test_regression.py       # 本次新增（golden 回归）
├── requirements.txt
├── requirements-dev.txt
└── pytest.ini
```

---

## 3. 核心架构变更

### 3.1 去 Windows COM 依赖

- 删除 `service/converter.py` 中 39+ 处 COM 调用（`DispatchEx` / `Documents.Open` / `Fields.Update` / `Quit`）
- 删除「杀 Word 进程」「轮询等文件锁释放」两个对抗性补丁
- 目录插入、分节、合并全部改为 OOXML 结构操作，不依赖 Word

### 3.2 去 Pandoc 链路

- 移除 pandoc 转换步骤（旧实现：pandoc 制造格式问题，再由后续 step 修复）
- 正文格式规范化改为直接操作 `document.xml`

### 3.3 排版引擎抽象

```text
DocConverter (ABC)
├── LibreOfficeConverter     Linux 默认实现，走 soffice + UNO
├── NoopConverter            无引擎降级（测试/纯 XML 场景）
└── (AsposeConverter)        预留逃生出口，见 REFACTOR_PLAN §3.3
```

- 业务代码只依赖 `DocConverter`，不依赖具体引擎
- 引擎缺失时**显式降级并告警**，不静默当作成功

### 3.4 20 步可枚举流水线

`PIPELINE_STEPS`（顺序即执行顺序）：

```
open → repair → numbering → orientation → footer → highlight → template
→ toc_cleanup → headings → split → cover → content → merge → toc_insert
→ toc_section → picture → finalize → engine → save → validate
```

- 每步一个方法，可独立计时、可独立失败
- 失败即抛 `AppError`，由流水线统一捕获并记录到步骤报告，**不允许静默交付坏文档**

---

## 4. 本次封面模板整合（专项）

### 4.1 背景

旧实现按 `文档类型 × 日期型 × 长度档` 维护 **24 份模板 + 7 份配置**：

```
reference_cover{document_type}_{date_type}_{length_param}.docx
cover_replace_config_{date_type}_{length_param}.json
```

逐份 diff 后发现 24 份模板只差三处：

| 维度 | 差异本质 |
|------|----------|
| 文档类型（1/2/3/4） | 只是文本（初步设计 / 可行性研究 / （代可研）/（代初设）） |
| 日期型（年月日/年月） | 只是文本 |
| 长度档（1/2/3 行） | 只差 1 个空段落 |

### 4.2 整合结果

收敛为 **1 份模板 + 1 份配置**，差异全部运行时回填：

```
backend/resources/cover_template.docx   ← 唯一封面模板
backend/config/cover_config.json        ← 文档类型/字段占位/元数据规则/格式
template/archive/                       ← 24 份旧模板 + 7 份旧配置归档
```

- 模板选择逻辑整体删除，改为「单一模板 + 运行时回填」
- 阶段副行（代可研/代初设）在类型 1/2 时整段删除，3/4 时回填
- 新增 `use_cover_template` 选项：单元测试用合成文档时走就地清理，不强制替换模板

### 4.3 顺带修复的两个 bug

1. **人员识别空格 bug**：源稿写「总 设 计 师：曹 永」，职务带空格导致关键词匹配失败，
   第三位人员一直留空。现职务/姓名先去空格再匹配，封面三人全部正确回填。
2. **悬空样式 bug**：封面包与正文包样式表不同，合并后正文引用的 `a/a0/a1/Normal`
   等样式悬空，`validate` 一直失败。在 `DocumentMerger` 增加 `_merge_styles`，
   合并时把源样式并入目标 styles.xml，`validate` 通过。

---

## 5. 测试与回归

### 5.1 测试套件

- 33 个测试全部通过
- 新增 `tests/test_regression.py`：`test/src.docx → obj.docx` 结构级对比
  （样例缺失时自动跳过）

### 5.2 回归结果

| 检查项 | 结果 |
|--------|------|
| 表格数量 | ✅ 20 = 20 |
| 分节数 | ✅ 6 = 6 |
| 目录域存在 | ✅ 1 = 1 |
| 高亮已清除 | ✅ 0 处 |
| 图片资源 | ✅ 保留 |
| 封面内容（项目名/日期/库号/人员） | ✅ 对齐 |

### 5.3 已知差异（与封面整合无关）

| 差异 | 说明 | 状态 |
|------|------|------|
| 正文文本保留度 85.9% | 实测为**测量口径偏差**：obj 的目录被 Word 展开成 140 条带页码条目，输出只有 1 条未刷新的域占位；剔除目录区后正文相似度 = 100% | 待接 LibreOffice 后自然消除 |
| 横向节数 1 vs 2 | 旧 COM 流程在横/竖切分时额外生成一个横向节 | 业务方确认系 COM 遗留 bug，无需处理 |

---

## 6. 常量外置清单

所有可调参数已迁入 `backend/config/*.json`（代码中禁止硬编码）：

| 文件 | 内容 |
|------|------|
| `cover_config.json` | 封面模板名、文档类型文字、字段占位、元数据抽取规则、字间距 |
| `layout_rules.json` | 表格字体、目录标题关键词、标题层级 |
| `document_settings.json` | 纸张尺寸、页边距、单位换算 |
| `picture_settings.json` | 图片对齐、行距、环绕方式 |

---

## 7. 待办事项

- [ ] 接入 LibreOffice 引擎刷新目录页码（需先安装 LibreOffice，见 §5.3）
- [ ] 确定封面「档案号/图号」的生成规则（当前为模板静态值）
- [ ] `sqls/` 目录去留（涉及历史表结构，待业务确认）
- [ ] 前端（Vite + React）骨架搭建
