"""LibreOffice 域刷新（目录页码 / 交叉引用）。

必须用 **LibreOffice 自带的 Python** 运行（它内置 `uno` 模块）：

    <LibreOffice>/program/python.exe refresh_fields.py --port <端口> <docx路径>

conda / 系统 Python 没有 `uno`，无法运行此脚本。

## 为什么本脚本只"连接"、不"启动"

实测结论（Windows + LibreOffice 26.8）：

- 由 LibreOffice 自带 Python 用 subprocess 启动 `soffice.exe --accept=...`
  **会失败**：进程起来了但不注册 UNO 监听端口，客户端永远连不上。
- 由普通进程（后端 Python / PowerShell）启动同一命令**正常监听**。

因此职责拆分：

| 角色 | 负责 |
|------|------|
| 后端 `LibreOfficeConverter` | 启动 soffice、等待端口就绪、结束时清理 |
| 本脚本（LO Python） | 只连接已就绪的端口，刷新索引并保存 |

## UNO 端点写法的坑

- **监听端**（`--accept=`）必须是 `;urp;StarOffice.ServiceManager`
- **连接端**（resolver）必须是 `;urp;StarOffice.ComponentContext`

两者写反会导致连接被拒绝，这是最容易踩的坑。

退出码：0 成功；2 参数错误；3 文件不存在；4 无法连接。
"""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path


def _parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="用 UNO 刷新 docx 的目录与域")
    parser.add_argument("document", help="待刷新的 docx 路径")
    parser.add_argument("--port", type=int, required=True, help="soffice UNO 监听端口")
    parser.add_argument("--timeout", type=float, default=60.0, help="连接与刷新超时（秒）")
    return parser.parse_args(argv)


def _connect(port: int, timeout: float):
    """连接已由外部启动的 soffice 实例。"""
    import uno  # noqa: PLC0415  仅 LibreOffice Python 可用

    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx
    )
    url = (
        f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
    )

    deadline = time.monotonic() + timeout
    last_error: Exception | None = None
    while time.monotonic() < deadline:
        try:
            return resolver.resolve(url)
        except Exception as exc:  # noqa: BLE001  UNO 连接失败异常类型不固定
            last_error = exc
            time.sleep(0.3)
    raise ConnectionError(f"无法连接 LibreOffice UNO 服务 (port={port}): {last_error}")


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv or sys.argv[1:])

    target = Path(args.document).resolve()
    if not target.exists():
        print(f"文件不存在: {target}", file=sys.stderr)
        return 3

    import uno  # noqa: PLC0415
    from com.sun.star.beans import PropertyValue  # noqa: PLC0415

    try:
        ctx = _connect(args.port, args.timeout)
    except ConnectionError as exc:
        print(str(exc), file=sys.stderr)
        return 4

    def prop(name: str, value) -> "PropertyValue":
        item = PropertyValue()
        item.Name = name
        item.Value = value
        return item

    desktop = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", ctx
    )

    url = uno.systemPathToFileUrl(str(target))
    document = desktop.loadComponentFromURL(
        url, "_blank", 0, (prop("Hidden", True),)
    )

    try:
        # 1) 刷新所有索引（目录、图表目录）
        indexes = document.getDocumentIndexes()
        count = indexes.getCount()
        for i in range(count):
            indexes.getByIndex(i).update()

        # 2) 刷新普通域（交叉引用、PAGEREF 等）
        try:
            document.getTextFields().refresh()
        except Exception:  # noqa: BLE001  部分文档无域
            pass

        document.store()
        print(f"已刷新 {count} 个索引并保存: {target.name}")
        return 0
    finally:
        try:
            document.close(True)
        except Exception:  # noqa: BLE001
            pass
        # 由本脚本负责关闭 soffice，避免后端还要处理进程树
        try:
            desktop.terminate()
        except Exception:  # noqa: BLE001
            pass


if __name__ == "__main__":
    sys.exit(main())
