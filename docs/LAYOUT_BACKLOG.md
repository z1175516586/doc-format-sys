# 排版待改清单

> 记录时间：2026-10-28
> 用途：改排版逻辑前先读本文。每条都写了「现在的代码在做什么」和「要改成什么」。
> 样式来源以这两个文件为准，不要再回仓库根 `template/`（该目录已删除）：
>
> - 正文与页眉：`backend/resources/content_template.docx`
> - 封面：`backend/resources/cover_template.docx`

改的时候以模板里的 OOXML 为准（`w:pStyle`、`w:rFonts`、`w:sz`、页眉部件、绘图关系），不要在代码里再写一套字号和字体。

---

## 1. 正文样式与页眉图标跟正文模板走

**要达到的效果**

目标文档的标题、正文样式与 `content_template.docx` 一致。模板里「标题一」是什么字体、字号、段前段后，输出就是什么。模板页眉左上角的公司图标，输出的页眉里也要有。

**现在的代码**

- 流水线 `_step_template`（`backend/app/services/format_service.py`）只在调用方传入 `template_path`，或 `backend/templates` 里注册了默认模板时，才调用 `TemplateService.apply_header_footer`。
- **没有**把 `backend/resources/content_template.docx` 当作内置正文模板。
- `HeadingProcessor`（`backend/app/domain/ooxml/headings.py`）只给段落套 `Heading 1..4` 的样式名，不把模板 `styles.xml` 里该样式的定义拷过来。源文档里如果没有这些样式的字体定义，标题看起来就还是正文。
- `apply_header_footer`（`backend/app/services/template_service.py`）按序号替换页眉/页脚部件的 XML。图标在 `w:drawing` / `w:blip` 里，靠 `r:embed` 指向 `word/media`。只换 XML、不把模板的图片部件和关系一起迁过去，图标就会丢。这和下面第 2 条是同一类问题。

**改的时候**

1. 默认正文模板固定为 `backend/resources/content_template.docx`，调用方未另选模板时用它。
2. 把模板中实际用到的段落样式定义合并进目标 `styles.xml`（按样式名对齐，而不是只写 styleId）。
3. 移植页眉时一并复制被引用的图片部件，并重写 `rId`。

---

## 2. 封面左上角公司图标丢失

**要达到的效果**

`cover_template.docx` 前两页左上角的公司图标，转换结果里仍然在。

**现在的代码**

`CoverService` 用 `backend/resources/cover_template.docx` 做底，再回填文字、移植设计人员表（`CoverProcessor.transplant_table_after_marker`）。文字回填本身不会删图，但有两处会把图弄丢：

- 封面模板里的图标若在图形 / 文本框中，`CoverProcessor._text_nodes` 会跳过 `w:drawing`、`w:pict`、`w:txbxContent`。这是为了避免把图形里的坐标数字拼进「库号」文本（`cover.py` 里已有这段注释）。跳过拼接是对的，**不能**在回填时把这些节点清空。
- 若之后用 LibreOffice `document.store()` 整包重存，曾经出现 `word/media` 被清空（见 `CHANGELOG.md` 的 P0）。在图标问题修好之前，不要靠整包重存来「刷新」。

**改的时候**

以封面模板为底包输出，确认 `word/media` 与 `document.xml.rels` 里的图片关系仍在。不要用「删掉 drawing 再重写段落」的方式填库号。

---

## 3. 页眉：左项目名称、右文档类型，横线位置跟正文模板

**要达到的效果**

- 页眉左侧：项目名称
- 页眉右侧：文档类型（初步设计 / 可行性研究 / 初步设计（代可研）/ 可行性研究（代初设），与 `cover_config.json` 的 `document_types` 一致）
- 页眉下方横线的位置、粗细与 `content_template.docx` 的页眉一致

**现在的代码**

页眉是整段从模板拷贝，没有按「项目名称 + 文档类型」填占位。项目名称在 `CoverMetadataExtractor` 里已经抽到了，文档类型在 `ConvertOptions.document_type` 里，这两处还没有接到页眉。

**改的时候**

先看正文模板页眉的结构（多半是一个两列的表格或两个制表位，横线是段落 `w:pBdr` 或表格底边）。保留这条横线的 OOXML，只替换左右两格的文字。不要自己用另一套边框数值重画。

---

## 4. 目录字体

**要达到的效果**

| 位置 | 字体 | 字号 | 其他 |
|---|---|---|---|
| 目录条目正文（标题文字） | 宋体 | 四号（14 磅） | |
| 「目录」二字 | 宋体 | 四号 | 加粗 |
| 引导点（……）和页码 | Times New Roman | 四号 | |

**现在的代码**

`TocProcessor._build_title_paragraph`（`backend/app/domain/ooxml/toc.py`）只给「目录」加了居中和 `w:b`，没有设宋体四号。目录条目是 TOC 域，刷新之后的字体来自 `TOC1` / `TOC2` / `TOC3` 样式，以及域里点号、页码各自的 run。代码没有改这些样式。

**改的时候**

1. 「目录」标题段落：宋体、四号、加粗。
2. 在 `styles.xml` 里把模板（或此处）的 TOC 样式设为宋体四号。
3. 点号和页码在 Word 里通常是单独的 run（或 `w:tab` 前导符）。刷新后检查这些 run 的 `w:rFonts` 的 ascii/hAnsi 是否为 Times New Roman，eastAsia 是否仍为宋体。只改点号和页码，不要把标题文字也改成 Times New Roman。

字号四号 = 28 半磅（`w:sz` 的值是半磅）。这个换算已有 `pt_to_half_points`，不要再写死魔法数散落各处；若要可配置，放进 `backend/config`。

---

## 5. 封面：库号字体、第二页图标、段落样式一一对应、项目名称换行

**要达到的效果**

- 第一页右下角「库号：」为黑体三号；冒号后面的编号为 Times New Roman 三号。
- 第二页除左上角图标（见第 2 条）外，现有结果可以保持。
- 封面模板里每个段落的样式都要保留，输出与模板段落一一对应，不要另起一套样式。
- 有人写项目名称时中间按了回车，分成两行。模板里项目名称通常是一个段落，两行要能收成模板需要的形式（或按模板的换行规则放进去）。

**现在的代码**

- 库号用整段字符串替换（`CoverProcessor.replace_fields`），替换时把多个 `w:t` 收成第一个 run，**字体会跟第一个 run 走**，做不到「库号：」和编号两种字体。
- 项目名称：`CoverMetadataExtractor._prepend_company` 只会在前一段像单位名（以公司/集团/厂等结尾）时，把单位名和项目名拼起来。**它不处理项目名称自身被回车拆成两行**，也不知道哪两行属于同一个名称。
- 封面阶段文字（初步设计等）会改写或删除模板段落，其余占位是按模板原文查找替换。模板段落的 `w:pStyle` 一般还在，但库号、项目名这类跨 run 替换可能把 run 级字体冲掉。

**改的时候**

1. 「库号：」和编号分成两个 run，分别设黑体三号、Times New Roman 三号。三号 = 16 磅 = `w:sz` 32。
2. 回填时只改 `w:t` 文本，保留该段落的 `w:pPr` / `w:pStyle`。
3. 项目名称换行：在封面区、设计人员名单之前，识别连续的、以「项目」收尾的多行（中间那行不以「项目」结尾）。如何合并要跟模板段落数对齐——模板只有一个项目名称段落时，两行合成一段再写入；不要把单位名和项目名的规则混在一起。

---

## 6. 表格：保留原字体和框线，只改行距和对齐

**要达到的效果**

表格内容、字体、框线与原文档一致。额外只做两件事：表格段落单倍行距，表格在页面上居中。

**现在的代码（和这条相反）**

`_step_content` 对正文全部表格调用 `TableProcessor.apply_font_to_all_tables`，字体来自 `layout_rules.json` 的 `fonts.table_default`（宋体 14 磅）。封面第一张表还会被改成 `table_cover`（宋体 9 磅），并可能被 `shrink_first_table_rows` 压行高。框线本身没有重画，但字体会被盖掉。

**改的时候**

正文表格不要再统一刷字体。只给表格所在段落设：

- 段落对齐 `w:jc` = `center`（若「表格居中」是指表本身，则是 `w:tblPr/w:jc`，不是单元格里的字）
- 行距单倍：`w:spacing` 的 `w:line` = `240`、`w:lineRule` = `auto`

单元格里的字和 `w:tblBorders` 保持不动。封面设计人员表若也要遵守这条，不要再套 `table_cover` 字体；行高压缩是为了避免封面溢出，改之前先看封面还会不会被撑破。

---

## 7. 「1 总论」丢失，以及不要用这句话来切分封面和正文

**现象**

测试文档转换后，标题「1 总论」没有了。旧实现用「1 总论」这类文字判断目录并切开封面和正文，这个判据不可靠。

**现在的代码**

切分在 `DocumentSplitter`（`backend/app/domain/ooxml/document.py`）：

1. 用目录关键词（`layout_rules.json` 的 `toc.keywords`，精确等于「目录」等）找目录段落。
2. 从目录之后找第一个带 Heading 样式的段落，当作正文起点，起点之前划给封面。
3. 找不到标题样式时，退路是「第一个以数字开头的段落」。

标题样式来自 `HeadingProcessor`：用编号正则把「1 总论」标成 Heading 1。顺序校验失败时会把标题改回正文，于是「1 总论」既可能没被标成标题，也可能被当成目录条目删掉。

`TocProcessor.remove_toc_blocks` 会拆掉旧的目录内容控件。若「1 总论」当时还在目录控件里（带页码的目录行，或控件边界划错），会被一起清掉，正文里的真标题就没了。

**改的时候**

用结构判断，不用某句固定标题：

- 封面区：从文档开头到「设计人员名单」标记（`cover_config.json` 的 `designers_marker`）及其后的设计人员表。标记必须存在才切，不存在就不要删表（`remove_tables_before_marker` 已有这条约束）。
- 目录区：`w:sdt` 且别名为 TOC，或段落内含 TOC 域（`w:instrText` 含 `TOC`）。只删这块，不要按「以数字开头」删段落。
- 正文起点：目录结构之后的第一块正文，而不是字符串「1 总论」。
- 编号标题仍要标成标题样式（否则新插入的 TOC 域是空的），但判据是编号层级，不是某一个标题的文字。标样式之前先把目录控件移走，避免目录行「1 总论 …… 1」被当成正文标题。

---

## 8. 现有实现里建议一起改的点

这些不是上面 7 条的原文，但是同一轮改样式时会碰到。

| 点 | 说明 |
|---|---|
| 引擎整包重存丢图 | LibreOffice `store()` 曾把 `word/media` 清掉。页眉和封面图标都依赖这些部件。刷新目录应只回写 TOC 区域，或在重存后把媒体部件合并回去。 |
| 库号串进图形坐标 | `CoverProcessor._replace_in_paragraph` 若把 drawing 里的 `w:t` 算进同一段，会出现「2753995…库号：…」。库号替换必须限定在普通 run。 |
| 封面模板查找还有旧目录 | `CoverService` 仍会到仓库根 `template/` 找 `cover_template.docx`。该目录已删除，实际命中的是 `backend/resources`。旧查找可以去掉，避免以后把文件放错位置还以为生效。 |
| 回归测试不再套页眉模板 | `tests/test_regression.py` 的 `template_path` 已是 `None`。正文模板改成默认读取 `content_template.docx` 之后，回归结果会变，断言要按新模板重看，不要为了维持旧断言把模板又绕开。 |

---

## 建议的改动顺序

1. 先定切分（第 7 条）。切错了，后面的封面、目录、正文样式都会标到错误的段落上。
2. 再保留封面图标和库号字体（第 2、5 条），确认媒体部件还在。
3. 接入 `content_template.docx` 的样式和页眉（第 1、3 条）。
4. 目录字体（第 4 条）放在 TOC 域插入并刷新之后核对。
5. 最后改表格（第 6 条）：去掉统一刷字体，只加单倍行距和表居中。
