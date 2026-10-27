# 目录页码刷新排障记录

> 记录时间：2026-10-27 23:20:06
> 主题：LibreOffice 引擎接入 + 目录页码刷新 + 回归差异定位
> 结论：目录页码已可自动刷新（140 条带页码），但引出图片丢失的新问题（见 §8）

---

## 1. 问题起点

回归测试中「正文文本保留度」只有 **85.9%**（期望 ≥ 98%）。
当时的判断是：obj.docx 的目录已被 Word 刷新成 140 条带页码条目，
而我们的输出只有 1 条未求值的 TOC 域占位，所以字符量对不上。

拆解验证（剔除目录区后）：

| 项 | obj.docx | 输出 |
|---|---|---|
| 正文总段数 | 859 | 730 |
| 其中目录条目 | 140 条 | 1 条域占位 |
| 剔除目录后 | 732 | 729 |
| 剔除目录后文本相似度 | — | **100.0%** |

**结论**：正文一字未丢，差异全部来自「目录未刷新」。要彻底解决，必须接排版引擎。

---

## 2. 第一层：引擎根本没装

```powershell
where.exe soffice        # 找不到
where.exe libreoffice    # 找不到
```

常见安装路径（`C:\Program Files\LibreOffice\program\soffice.exe` 等）也全部不存在。

**动作**：`winget install --id TheDocumentFoundation.LibreOffice -e --silent`
安装成功（LibreOffice 26.8.1.1）。

---

## 3. 第二层：bootstrap.ini「已损坏」

安装完成后启动报错：

> 无法启动应用程序。
> 配置文件「C:\Program Files\LibreOffice\program\bootstrap.ini」已经损坏

### 3.1 第一次误判：占位符未替换

读文件内容：

```ini
[Bootstrap]
InstallMode=<installmode>          ← 占位符没被替换
ProductKey=LibreOffice 26.8
UserInstallation=$SYSUSERCONFIG/LibreOffice/4
```

`<installmode>` 是 MSI 安装时的模板变量，winget 静默安装没填。
用管理员权限改成 `InstallMode=install`。

**结果：仍然报同样的错。**

### 3.2 真正的根因：UTF-8 BOM

检查文件原始字节：

```
EF BB BF 5B 42 6F 6F 74 73 74 72 61 70 5D ...
└─ BOM ─┘ [  B  o  o  t  s  t  r  a  p  ]
```

**首三个字节是 UTF-8 BOM（`EF BB BF`）**。

原因是修复时用了 PowerShell 的：

```powershell
Set-Content -Path $p -Value $c2 -Encoding UTF8
```

Windows PowerShell 5.1 的 `-Encoding UTF8` 会**自动写入 BOM**。
而 LibreOffice 的 INI 解析器要求首字节就是 `[`，读到 BOM 直接判定文件损坏。

### 3.3 正确修法

用字节写入，指定纯 ASCII + CRLF（Windows INI 惯例）：

```powershell
$lines = @(
  "[Bootstrap]"
  "InstallMode=install"
  "ProductKey=LibreOffice 26.8"
  "UserInstallation=`$SYSUSERCONFIG/LibreOffice/4"
)
$content = ($lines -join "`r`n") + "`r`n"
[System.IO.File]::WriteAllBytes($p, [System.Text.Encoding]::ASCII.GetBytes($content))
```

验证字节：`5B 42 6F 6F 74 73`（`[Boots`，无 BOM）。

### 3.4 验证方式本身也有坑

用 `soffice.exe --version` 验证是**无效的**：

- GUI 程序不向控制台输出 `--version`
- 命令会一直挂起，看起来像"又坏了"

正确的验证方式是做一次真实转换：

```powershell
soffice --headless --convert-to txt --outdir <dir> <file.docx>
```

本次验证结果：`exit=0`，产出 86,738 字节的 `obj.txt` → 引擎可用。

---

## 4. 第三层：UNO 端口始终连不上

引擎能转换，但域刷新（需要 UNO）失败。逐步排除：

### 4.1 排除「uno 模块缺失」

conda 的 Python 确实没有 `uno`（`ModuleNotFoundError`）。
改用 LibreOffice 自带的 Python：

```powershell
& "C:\Program Files\LibreOffice\program\python.exe" -c "import uno"
# uno import OK  C:\Program Files\LibreOffice\program\uno.py
```

→ `uno` 可用。

### 4.2 排除「soffice 路径探测错误」

LibreOffice Python 的 `sys.executable` 是 `...\program\python-core-3.13.15`，
但 `Path(sys.executable).resolve().parent` 正确解析到 `...\program`，
`soffice.exe` 判定存在 → 探测逻辑没问题。

### 4.3 真正的根因 A：`--accept` 的服务名写反

启动时检查端口监听：

```powershell
Get-NetTCPConnection -LocalPort 2002
# 空 —— 没有任何监听
```

**问题在 `--accept` 参数的结尾**：

| 用途 | 正确写法 |
|------|----------|
| **监听端**（`--accept=`） | `socket,host=...,port=...;urp;StarOffice.ServiceManager` |
| **连接端**（resolver URL） | `socket,host=...,port=...;urp;StarOffice.ComponentContext` |

两者**不能混用**。原代码把监听端也写成了 `ComponentContext`，
soffice 因此不注册监听服务，客户端永远连不上。

### 4.4 真正的根因 B：谁来启动 soffice

这是最关键的一条，实测得出的反直觉结论：

| 启动者 | 结果 |
|--------|------|
| 由 **LibreOffice 自带 Python** 用 `subprocess.Popen` 启动 | ❌ 进程起来了，但**端口不监听** |
| 由 **普通进程**（PowerShell / 后端 Python）启动同一命令 | ✅ 端口正常监听 |

验证方式：
- 失败时：`soffice` / `soffice.bin` 进程都在，`Get-NetTCPConnection` 无 2002 端口
- 成功时：端口 `127.0.0.1:2002 Listen`

**因此职责必须拆开**：

```
后端 Python (LibreOfficeConverter)
   ├─ 1. 启动 soffice --accept=...  并轮询等待端口就绪
   ├─ 2. 调用 LO Python 执行 refresh_fields.py --port <port>
   └─ 3. taskkill /F /T 清理进程树 + 删 profile 目录

LO Python (scripts/refresh_fields.py)
   └─ 只做：连接端口 → 刷新索引 → store() 保存 → 关闭
```

---

## 5. 第四层：刷新成功了，但我一开始验证错了

UNO 刷新返回 `REFRESH OK`、`indexes count: 1`，但用
`pkg.paragraphs()` 找 TOC 条目却是 0 —— 差点误判为"刷新无效"。

**原因**：`pkg.paragraphs()` 只返回 `body` 的**顶层**段落，
而 TOC 条目被包在 `w:sdt`（结构化文档标签）里，是嵌套结构。

正确遍历方式：

```python
for para in pkg.body.iter(qn("w:p")):   # 递归
    ...
```

用递归后统计：

```
TOC1: 55 条   TOC2: 72 条   TOC3: 13 条   合计 140 条
1.1项目背景 …… 1
1.4设计原则 …… 2
13.3结论 …… 56
```

**与 obj.docx 的 140 条完全对得上**，页码也正确。

---

## 6. 第五层：流水线顺序把刷新结果覆盖了

单测刷新成功、端到端却还是空目录。检查 `PIPELINE_STEPS`：

```
... → picture → finalize → engine → save → validate
                                    ↑
                          在 engine 之后又写了一遍内存包
```

`_step_engine` 是**就地修改文件**（LibreOffice `store()` 回写同一个 docx），
但紧跟着的 `_step_save` 又把内存里那份**未刷新**的 `DocxPackage` 写回文件，
刷新结果被完整覆盖。

**修正**：

```python
"picture",
"finalize",
# save 必须在 engine 之前：engine 会就地刷新并回写文件，
# 若 save 在其后再写一次内存包，会把刷新结果覆盖掉。
"save",
"engine",
"validate",
```

修正后端到端验证通过：`engine: ok=True stats={'refreshed': True}`，目录 140 条带页码。

---

## 7. 排障中踩过的无效路径

记录下来避免重走：

| 尝试 | 结果 | 为何无效 |
|------|------|----------|
| 注入 `<w:updateFields w:val="true"/>` 后 `--convert-to docx` | ❌ 目录仍是占位 | LO 的 `--convert-to` 不做分页重算，忽略该标志 |
| `soffice.exe --version` 验证安装 | ❌ 挂起 | GUI 程序不输出该信息，会阻塞 |
| 把 `InstallMode=<installmode>` 改成 `install` 就重启 | ❌ 仍报损坏 | 真正的病根是 BOM，占位符只是表象 |
| 用 `pkg.paragraphs()` 找 TOC | ❌ 误判为 0 条 | TOC 在 `w:sdt` 内，需递归遍历 |
| LO Python 自行 `Popen` 启动 soffice | ❌ 端口不监听 | 见 §4.4，必须由外部进程启动 |

---

## 8. 遗留问题（本次新发现，未解决）

接入引擎、修正顺序后，媒体文件数从 3 变成 **0**：

| 项 | 输出（引擎刷新后） | obj.docx |
|---|---|---|
| media 部件 | **0** | 2 |
| 段落数 | 779 | 914 |
| TOC 条目 | 140 | 127 |

**判断**：`LibreOffice` 的 `document.store()` 重新打包 docx 时丢掉了图片部件。

**待办**：

- [ ] 定位 LO store 丢图的原因（可能是关系/内容类型不完整，或图片为 WMF 等旧格式）
- [ ] 若无法保留，改用「只把刷新后的 TOC 段落回填到原文档」的策略，
      不用 LO 整体保存，避免动到图片与其它部件

---

## 9. 最终架构（已落地）

### 9.1 引擎探测

```python
_resolve_binary()   # PATH → 常见安装路径（Program Files / usr/bin）
_resolve_python()   # soffice 同目录下的 python.exe；Linux 回退 /usr/bin/python3
```

`capabilities().refresh_fields` 由 `_resolve_python()` 决定，
找不到 LO Python 时如实声明"不支持"，不虚报。

### 9.2 刷新链路

```
LibreOfficeConverter.refresh_fields(path)
  ├─ _free_port()                       # 向系统申请空闲端口，避免并发冲突
  ├─ Popen(soffice --accept=...ServiceManager)  # 主进程启动
  ├─ _wait_for_port(port, 45s)          # 冷启动要初始化 profile，超时必须给足
  ├─ run(lo_python refresh_fields.py --port <port> <docx>)
  └─ _terminate_tree()                  # taskkill /F /T（soffice.bin 才是真进程）
     + rmtree(profile_dir)
```

### 9.3 相关文件

| 文件 | 变更 |
|------|------|
| `backend/app/converters/libreoffice.py` | 重写刷新链路；新增 `_resolve_python` / `_free_port` / `_wait_for_port` / `_terminate_tree` / `_find_in_common_locations` |
| `backend/scripts/refresh_fields.py` | 新建；改为「只连接」，支持 `--port` / `--timeout` |
| `backend/app/services/format_service.py` | `PIPELINE_STEPS` 调整：`save` 移到 `engine` 之前 |

---

## 10. 可复现的验证命令

```powershell
# 1. 确认引擎探测
cd backend
python -c "from app.converters.libreoffice import LibreOfficeConverter; print(LibreOfficeConverter().capabilities().to_dict())"

# 2. 端到端 + 回归对比
python _tmp_final.py    # 见历史脚本，或直接用 tests/test_regression.py

# 3. 检查 TOC 条目（必须递归遍历 body）
```

```python
from app.domain.ooxml.package import DocxPackage
from app.core.namespaces import qn

with DocxPackage.open("out.docx") as pkg:
    n = sum(
        1 for p in pkg.body.iter(qn("w:p"))
        if (p.find(f"{qn('w:pPr')}/{qn('w:pStyle')}") is not None)
    )
```

---

## 11. 一句话总结

目录页码不刷新，表面是"引擎没装"，实际是**五层问题叠加**：
引擎缺失 → BOM 破坏配置 → `--accept` 服务名写反 → 启动者选错 → 流水线顺序覆盖结果。
每一层单独看都像"配置坏了"，只有逐层用字节/端口/进程状态证伪才能定位。
