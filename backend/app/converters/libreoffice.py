"""LibreOffice 引擎实现。

关键点（也是原 COM 方案无法做到的）：
1. **进程隔离**：每个任务使用独立 UserInstallation profile。
   LibreOffice 默认共享 profile，并发时会直接报 "LibreOffice is already running"
   然后退出——这是 unoconv / JODConverter 都要解决的核心问题。
2. **超时保护**：soffice 偶发挂起，必须有硬超时，否则会拖垮 worker 池。
3. **跨平台**：Windows / Linux 同一份实现，仅二进制路径不同。

能力边界（诚实声明）：
- doc -> docx、docx -> pdf：由 soffice 命令行完成，可靠
- refresh_fields：需要 UNO 绑定；未安装 uno 模块时声明为不支持
"""

from __future__ import annotations

import os
import shutil
import socket
import subprocess
import tempfile
import time
import uuid
from pathlib import Path

from app.converters.base import ConverterCapabilities, DocConverter
from app.core.config import get_settings
from app.core.errors import EngineUnavailableError, LayoutError
from app.core.logging import get_logger

logger = get_logger(__name__)


def _uno_available() -> bool:
    try:
        import uno  # noqa: F401
    except Exception:
        return False
    return True


class LibreOfficeConverter(DocConverter):
    """基于 soffice 命令行的排版引擎。"""

    def __init__(self, binary: str | None = None, timeout_seconds: int | None = None) -> None:
        settings = get_settings()
        self._binary_setting = binary or settings.libreoffice_binary
        self.timeout = timeout_seconds or settings.libreoffice_timeout_seconds
        self._resolved: str | None = None

    # ------------------------------------------------------------------
    # 探测
    # ------------------------------------------------------------------
    def _resolve_binary(self) -> str | None:
        if self._resolved is not None:
            return self._resolved or None

        candidate = self._binary_setting
        if candidate and Path(candidate).is_file():
            self._resolved = candidate
            return candidate

        found = (
            shutil.which(candidate)
            or shutil.which("libreoffice")
            or shutil.which("soffice")
            or _find_in_common_locations()
        )
        self._resolved = found or ""
        return found

    def _resolve_python(self) -> str | None:
        """定位能运行 UNO 脚本的解释器（LibreOffice 自带 Python）。

        普通 Python（conda / venv / 系统）没有 uno 模块，这是 LibreOffice
        UNO 绑定的机制决定的：uno 由 LibreOffice 安装自带。
        """
        binary = self._resolve_binary()
        if not binary:
            return None
        program_dir = Path(binary).resolve().parent
        candidates = []
        if os.name == "nt":
            candidates.append(program_dir / "python.exe")
        else:
            candidates.append(program_dir / "python")
            # Debian/Ubuntu 装 python3-uno 后，系统 python3 可直接 import uno
            candidates.append(Path("/usr/bin/python3"))
        for candidate in candidates:
            if candidate.is_file():
                return str(candidate)
        return None

    def capabilities(self) -> ConverterCapabilities:
        binary = self._resolve_binary()
        if not binary:
            return ConverterCapabilities(
                name="libreoffice",
                available=False,
                notes=("未找到 soffice 可执行文件",),
            )

        refresh_ok = self._resolve_python() is not None or _uno_available()
        notes = []
        if not refresh_ok:
            notes.append("未找到 LibreOffice Python，目录页码不会自动刷新")
        notes.append(f"binary={binary}")

        return ConverterCapabilities(
            name="libreoffice",
            available=True,
            doc_to_docx=True,
            refresh_fields=refresh_ok,
            export_pdf=True,
            page_count=refresh_ok,
            render_page=False,
            notes=tuple(notes),
        )

    # ------------------------------------------------------------------
    # 转换
    # ------------------------------------------------------------------
    def convert(self, src: Path, dst: Path, target_format: str) -> Path:
        binary = self._resolve_binary()
        if not binary:
            raise EngineUnavailableError("未找到 LibreOffice（soffice），无法执行转换")

        src = Path(src)
        dst = Path(dst)
        if not src.exists():
            raise LayoutError(f"源文件不存在: {src}")

        dst.parent.mkdir(parents=True, exist_ok=True)
        profile_dir = Path(tempfile.gettempdir()) / f"lo-profile-{uuid.uuid4().hex}"

        cmd = [
            binary,
            "--headless",
            "--norestore",
            "--nolockcheck",
            "--nodefault",
            "--nofirststartwizard",
            f"-env:UserInstallation=file://{profile_dir.as_posix()}",
            "--convert-to",
            target_format,
            "--outdir",
            str(dst.parent),
            str(src),
        ]

        logger.info("LibreOffice 转换: %s -> %s", src.name, target_format)
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                timeout=self.timeout,
                check=False,
                creationflags=_creation_flags(),
            )
        except subprocess.TimeoutExpired as exc:
            raise LayoutError(
                f"LibreOffice 转换超时（>{self.timeout}s），文档可能过于复杂"
            ) from exc
        finally:
            shutil.rmtree(profile_dir, ignore_errors=True)

        produced = dst.parent / f"{src.stem}.{_extension_of(target_format)}"
        if not produced.exists():
            stderr = (result.stderr or b"").decode("utf-8", errors="ignore").strip()
            raise LayoutError(f"LibreOffice 未产出目标文件: {stderr or '无错误输出'}")

        if produced.resolve() != dst.resolve():
            shutil.move(str(produced), str(dst))
        return dst

    # ------------------------------------------------------------------
    # 域刷新（主进程启动 soffice，LO Python 只负责连接）
    # ------------------------------------------------------------------
    def refresh_fields(self, path: Path) -> bool:
        """刷新目录页码与域。

        实测（Windows + LO 26.8）：让 LibreOffice 自带 Python 启动 soffice
        会失败（进程起来但不注册监听端口）。因此由**本进程**启动 soffice，
        LO Python 只做连接与刷新，这是唯一稳定的组合。
        """
        lo_python = self._resolve_python()
        binary = self._resolve_binary()

        if lo_python and binary:
            return self._refresh_via_socket(binary, lo_python, Path(path))

        # 兜底：当前进程内已有 uno（Linux 装了 python3-uno 的情况）
        if _uno_available():
            return self._refresh_fields_via_uno(Path(path))

        logger.info("未找到 LibreOffice Python，跳过域刷新")
        return False

    def _refresh_via_socket(self, binary: str, lo_python: str, path: Path) -> bool:
        """启动 soffice 监听 → LO Python 连接刷新 → 清理。"""
        script = Path(__file__).resolve().parents[2] / "scripts" / "refresh_fields.py"
        if not script.is_file():
            logger.warning("域刷新脚本缺失: %s", script)
            return False

        port = _free_port()
        profile_dir = Path(tempfile.gettempdir()) / f"lo-profile-{uuid.uuid4().hex}"
        # 监听端必须是 ServiceManager；连接端用 ComponentContext（见脚本注释）
        accept = f"socket,host=127.0.0.1,port={port};urp;StarOffice.ServiceManager"

        process = subprocess.Popen(
            [
                binary,
                "--headless",
                "--norestore",
                "--nolockcheck",
                "--nodefault",
                "--nofirststartwizard",
                "--invisible",
                f"--accept={accept}",
                f"-env:UserInstallation=file:///{profile_dir.as_posix().lstrip('/')}",
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=_creation_flags(),
        )

        try:
            if not _wait_for_port(port, timeout=45.0):
                logger.warning("LibreOffice UNO 端口 %d 未就绪，跳过域刷新", port)
                return False

            result = subprocess.run(
                [lo_python, str(script), "--port", str(port), str(path.resolve())],
                capture_output=True,
                timeout=self.timeout,
                check=False,
                creationflags=_creation_flags(),
            )
        except subprocess.TimeoutExpired:
            logger.warning("域刷新超时（>%ss）", self.timeout)
            return False
        finally:
            _terminate_tree(process)
            shutil.rmtree(profile_dir, ignore_errors=True)

        if result.returncode != 0:
            stderr = (result.stderr or b"").decode("utf-8", errors="ignore").strip()
            logger.warning("域刷新失败: %s", stderr or f"退出码 {result.returncode}")
            return False

        stdout = (result.stdout or b"").decode("utf-8", errors="ignore").strip()
        logger.info("域刷新成功: %s", stdout)
        return True

    def _refresh_fields_via_uno(self, path: Path) -> bool:
        """通过 UNO 连接常驻 soffice 刷新索引与域。

        必须走 UNO 而不是 --convert-to：后者只做格式转换，
        不会重新计算分页，因此目录页码不会更新。
        """
        try:
            import uno  # noqa: F401
            from com.sun.star.beans import PropertyValue  # type: ignore
        except Exception as exc:
            logger.warning("UNO 绑定导入失败: %s", exc)
            return False

        binary = self._resolve_binary()
        if not binary:
            return False

        profile_dir = Path(tempfile.gettempdir()) / f"lo-profile-{uuid.uuid4().hex}"
        port = 2002 + (os.getpid() % 500)
        accept = f"socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"

        process = None
        try:
            process = subprocess.Popen(
                [
                    binary,
                    "--headless",
                    "--norestore",
                    "--invisible",
                    "--nologo",
                    f"-env:UserInstallation=file://{profile_dir.as_posix()}",
                    f"--accept={accept}",
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=_creation_flags(),
            )

            local_ctx = uno.getComponentContext()
            resolver = local_ctx.ServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", local_ctx
            )

            ctx = None
            for _ in range(40):
                try:
                    ctx = resolver.resolve(
                        f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
                    )
                    break
                except Exception:
                    import time

                    time.sleep(0.5)

            if ctx is None:
                logger.warning("无法连接 LibreOffice UNO 服务")
                return False

            desktop = ctx.ServiceManager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", ctx
            )

            def prop(name: str, value: str) -> "PropertyValue":
                item = PropertyValue()
                item.Name = name
                item.Value = value
                return item

            url = uno.systemPathToFileUrl(str(path.resolve()))
            document = desktop.loadComponentFromURL(
                url, "_blank", 0, (prop("Hidden", "true"),)
            )

            # 刷新所有索引（目录）与域
            indexes = document.getDocumentIndexes()
            for i in range(indexes.getCount()):
                indexes.getByIndex(i).update()

            try:
                document.getTextFields().refresh()
            except Exception:
                pass

            document.store()
            document.close(True)
            return True

        except Exception as exc:
            logger.warning("域刷新失败: %s", exc)
            return False
        finally:
            if process is not None:
                process.terminate()
                try:
                    process.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    process.kill()
            shutil.rmtree(profile_dir, ignore_errors=True)

    # ------------------------------------------------------------------
    # 其它能力
    # ------------------------------------------------------------------
    def page_count(self, path: Path) -> int | None:
        # 依赖 UNO 的分页计算；未安装 uno 时明确返回 None，不猜测
        if not _uno_available():
            return None
        return None

    def export_pdf(self, src: Path, dst: Path) -> Path | None:
        try:
            return self.convert(Path(src), Path(dst), "pdf")
        except (EngineUnavailableError, LayoutError) as exc:
            logger.warning("导出 PDF 失败: %s", exc)
            return None


def _free_port() -> int:
    """向系统申请一个空闲端口，避免多任务并发时端口冲突。"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def _wait_for_port(port: int, timeout: float = 45.0) -> bool:
    """轮询等待 soffice 的 UNO 监听端口就绪。

    soffice 首次启动要初始化 profile，冷启动可能超过 20 秒，
    因此超时必须给足，不能沿用命令转换的短超时。
    """
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(0.5)
            if sock.connect_ex(("127.0.0.1", port)) == 0:
                return True
        time.sleep(0.3)
    return False


def _terminate_tree(process: subprocess.Popen) -> None:
    """结束 soffice 进程树。

    Windows 上 soffice.exe 只是启动器，真正的进程是 soffice.bin，
    必须按进程树结束，否则会残留占用 profile 目录。
    """
    if process.poll() is not None:
        return

    if os.name == "nt":
        try:
            subprocess.run(
                ["taskkill", "/F", "/T", "/PID", str(process.pid)],
                capture_output=True,
                timeout=15,
                check=False,
                creationflags=_creation_flags(),
            )
            return
        except (OSError, subprocess.TimeoutExpired):
            pass

    process.terminate()
    try:
        process.wait(timeout=10)
    except subprocess.TimeoutExpired:
        process.kill()


def _extension_of(target_format: str) -> str:
    """把 LibreOffice 的目标格式名映射为扩展名。"""
    mapping = {
        "docx": "docx",
        "doc": "doc",
        "pdf": "pdf",
        "MS Word 2007 XML": "docx",
        "writer_pdf_Export": "pdf",
    }
    return mapping.get(target_format, target_format.split(":")[0])


def _find_in_common_locations() -> str | None:
    """在 PATH 之外探测 soffice 的常见安装位置。"""
    if os.name == "nt":
        program_files = os.environ.get("ProgramFiles", r"C:\Program Files")
        program_files_x86 = os.environ.get(
            "ProgramFiles(x86)", r"C:\Program Files (x86)"
        )
        candidates = (
            Path(program_files) / "LibreOffice" / "program" / "soffice.exe",
            Path(program_files_x86) / "LibreOffice" / "program" / "soffice.exe",
        )
    else:
        candidates = (
            Path("/usr/bin/soffice"),
            Path("/usr/bin/libreoffice"),
            Path("/usr/local/bin/soffice"),
        )
    for candidate in candidates:
        if candidate.is_file():
            return str(candidate)
    return None


def _creation_flags() -> int:
    """Windows 下隐藏控制台窗口，其它平台返回 0。"""
    if os.name == "nt":
        return getattr(subprocess, "CREATE_NO_WINDOW", 0)
    return 0
