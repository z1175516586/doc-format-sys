import json
import os
import sys
import tempfile
import traceback
import threading
import asyncio
from utils import doc2docx
from service.convert_history_service import ConvertHistoryService
from pathlib import Path
from service.converter import DocumentConverter

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from utils.check_history_logo import print_dochistoryid

from fastapi.responses import StreamingResponse

# 将项目根目录加入 Python 搜索路径
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

app = FastAPI(
    title="文档转换服务",
    description="将 DOCX 文档转换为标准格式的 Web 服务",
    version="1.0.0",
)

BASE_DIR = Path(__file__).resolve().parent.parent
STATIC_DIR = BASE_DIR / "static"
TEMPLATE_DIR = BASE_DIR / "template"
TEMPLATES_DIR = BASE_DIR / "templates"
RESULT_DIR = BASE_DIR / "result"
TEMP_DIR = BASE_DIR / "temp"
LOG_DIR = BASE_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

# 挂载静态文件目录，提供前端资源
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

# 配置模板目录
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))

# 创建必要的目录
RESULT_DIR.mkdir(exist_ok=True)
TEMP_DIR.mkdir(exist_ok=True)


@app.get("/")
async def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.get("/api")
async def root():
    return {"message": "文档转换服务正在运行", "version": "1.0.0"}


@app.get("/result/{history_id}")
async def download(history_id: str):

    file_path = ConvertHistoryService.get_file_path_by_hash(history_id)

    if not file_path:
        raise HTTPException(status_code=404, detail="文件不存在")

    return FileResponse(
        file_path,
        filename=os.path.basename(file_path),
        media_type="application/octet-stream"
    )





class DocumentConverterService:
    def __init__(
            self,
            file: UploadFile,
            template_file: UploadFile | None,
            header_text: str,
            header_left_text: str,
            header_right_text: str,
            toc_title: str,
            save_intermediate: bool,
            document_type: int,
    ) -> None:
        self.file = file
        self.template_file = template_file
        self.header_text = header_text
        self.header_left_text = header_left_text
        self.header_right_text = header_right_text
        self.toc_title = toc_title
        self.save_intermediate = save_intermediate
        self.document_type = document_type
        self.source_file_path: str | None = None
        self.template_file_path: str | None = None

    def _ensure_docx_format(self, file_path: str) -> str:
        """
        如果是 .doc 转为 .docx
        如果是 .docx 直接返回
        其他格式报错
        """
        ext = Path(file_path).suffix.lower()

        if ext == ".docx":
            return file_path

        if ext == ".doc":
            # 复用你给的转换逻辑
            doc2docx.convert_doc_to_docx(file_path)
            docx_path = str(Path(file_path).with_suffix(".docx"))

            if not os.path.exists(docx_path):
                raise HTTPException(status_code=500, detail="DOC 转 DOCX 失败")

            return docx_path

        raise HTTPException(status_code=400, detail="仅支持 DOC 或 DOCX 格式文件")

    async def prepare_files(self) -> None:
        # 保存上传的源文件到临时路径
        suffix = Path(self.file.filename).suffix
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as temp_file:
            content = await self.file.read()
            temp_file.write(content)
            self.source_file_path = temp_file.name

        # 🔥 这里新增
        self.source_file_path = self._ensure_docx_format(self.source_file_path)

        # 如果提供了模板文件，保存至临时路径
        if self.template_file and self.template_file.filename:
            if not self.template_file.filename.endswith(".docx"):
                raise HTTPException(status_code=400, detail="模板文件必须是 DOCX 格式")

            with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as temp_template_file:
                content = await self.template_file.read()
                temp_template_file.write(content)
                self.template_file_path = temp_template_file.name

    async def convert_document_stream(self):

        import queue
        import io

        log_queue = queue.Queue()
        done_event = threading.Event()
        result_holder = {"result": None}

        class QueueWriter(io.TextIOBase):
            def write(self, s):
                s = s.strip()
                if s:
                    log_queue.put(s)

        def run_conversion():

            old_stdout = sys.stdout
            sys.stdout = QueueWriter()

            try:

                import datetime
                import re
                import socket
                import uuid
                import shutil

                now = datetime.datetime.now()
                timestamp = now.strftime("%y%m%d%H%M%S")

                filename_part = Path(self.file.filename).stem
                parts = filename_part.split("_")
                last_part = parts[-1] if parts else filename_part

                if re.search(r"[\u4e00-\u9fff]", last_part):
                    safe_part = last_part
                else:
                    safe_part = ""

                output_filename = f"{timestamp}_{safe_part}.docx" if safe_part else f"{timestamp}.docx"
                output_file_path = RESULT_DIR / output_filename

                template_path = self.template_file_path
                if not template_path:
                    template_file = TEMPLATE_DIR / "reference_content.docx"
                    if not template_file.exists():
                        for template_name in ["reference.docx", "template.docx"]:
                            candidate = TEMPLATE_DIR / template_name
                            if candidate.exists():
                                template_file = candidate
                                break
                    template_path = str(template_file) if template_file.exists() else None

                print("开始检查历史记录")

                history_data = ConvertHistoryService.check_history(self.source_file_path)

                print("历史记录检查完成")

                with DocumentConverter(document_type=self.document_type) as converter:

                    success, msg = converter.convert_document(
                        source_file=self.source_file_path,
                        output_file=str(output_file_path),
                        template_file=template_path,
                        header_text=self.header_text,
                        header_left_text=self.header_left_text,
                        header_right_text=self.header_right_text,
                        toc_title=self.toc_title,
                        save_intermediate=self.save_intermediate,
                        document_type=self.document_type,
                    )

                user_ip = None
                user_mac = None

                try:
                    user_ip = socket.gethostbyname(socket.gethostname())
                except:
                    pass

                try:
                    user_mac = ":".join(["%02x" % ((uuid.getnode() >> ele) & 0xff)
                                         for ele in range(0, 8 * 6, 8)][::-1])
                except:
                    pass

                history_id = ConvertHistoryService.update_history(
                    docx_path=str(output_file_path),
                    user_ip=user_ip,
                    user_mac=user_mac,
                    success=1 if success else 0,
                    msg=msg,
                    hash_code=history_data["history_id"],
                )

                result_holder["result"] = {
                    "success": success,
                    "filename": output_filename,
                    "history_id": history_id,
                    "message": msg
                }

            except Exception as e:

                result_holder["result"] = {
                    "success": False,
                    "message": str(e)
                }

            finally:

                sys.stdout = old_stdout

                if self.source_file_path and os.path.exists(self.source_file_path):
                    os.unlink(self.source_file_path)

                if self.template_file_path and os.path.exists(self.template_file_path):
                    os.unlink(self.template_file_path)

                done_event.set()

        thread = threading.Thread(target=run_conversion, daemon=True)
        thread.start()

        while not done_event.is_set():

            try:
                log = log_queue.get_nowait()

                yield f"data: {json.dumps({'type': 'log', 'message': log})}\n\n"

            except:
                await asyncio.sleep(0.1)

        result = result_holder["result"]

        if result and result["success"]:

            yield f"data: {json.dumps({'type': 'complete', 'data': result})}\n\n"

        else:

            yield f"data: {json.dumps({'type': 'error', 'data': result})}\n\n"

    async def convert_document(self) -> dict:
        """线程化执行 DocumentConverter，保持接口 async 简洁"""
        import socket
        import uuid
        import shutil
        result_holder = {"result": None}
        done_event = threading.Event()

        def run_conversion():
            try:
                import datetime
                import re

                # 1️⃣ 生成时间戳
                now = datetime.datetime.now()
                timestamp = now.strftime("%y%m%d%H%M%S")  # YYMMDDHHMMSS

                # 2️⃣ 处理原文件名
                filename_part = Path(self.file.filename).stem  # 去掉后缀
                parts = filename_part.split("_")
                last_part = parts[-1] if parts else filename_part

                # 3️⃣ 检查最后一节是否包含中文
                if re.search(r"[\u4e00-\u9fff]", last_part):
                    safe_part = last_part
                else:
                    safe_part = ""  # 不包含中文则只用时间戳

                # 4️⃣ 组合最终文件名
                output_filename = f"{timestamp}_{safe_part}.docx" if safe_part else f"{timestamp}.docx"
                output_file_path = RESULT_DIR / output_filename

                # 5️⃣ 处理模板文件路径
                template_path = self.template_file_path
                if not template_path:
                    template_file = TEMPLATE_DIR / "reference_content.docx"
                    if not template_file.exists():
                        for template_name in ["reference.docx", "template.docx"]:
                            candidate = TEMPLATE_DIR / template_name
                            if candidate.exists():
                                template_file = candidate
                                break
                    template_path = str(template_file) if template_file.exists() else None
                print("=======================")
                print_dochistoryid(self.source_file_path)
                history_data = ConvertHistoryService.check_history(self.source_file_path)
                print(history_data)
                print("=======================")
                # 6️⃣ 调用 DocumentConverter
                with DocumentConverter(document_type=self.document_type) as converter:
                    success, msg = converter.convert_document(
                        source_file=self.source_file_path,
                        output_file=str(output_file_path),
                        template_file=template_path,
                        header_text=self.header_text,
                        header_left_text=self.header_left_text,
                        header_right_text=self.header_right_text,
                        toc_title=self.toc_title,
                        save_intermediate=self.save_intermediate,
                        document_type=self.document_type,
                    )
                    # #
                    # success = 1
                    # msg = "ok"
                    # # 🔥 测试用，把源文件复制到目标路径
                    # import shutil
                    # shutil.copyfile(self.source_file_path, output_file_path)

                # 7️⃣ 获取用户 IP / MAC（示例：服务器本机，可从前端传入）
                user_ip = None
                user_mac = None
                try:
                    user_ip = socket.gethostbyname(socket.gethostname())
                except Exception:
                    pass
                try:
                    user_mac = ":".join(["%02x" % ((uuid.getnode() >> ele) & 0xff)
                                         for ele in range(0, 8 * 6, 8)][::-1])
                except Exception:
                    pass

                # 8️⃣ 成功或失败都调用 update_history
                if not success:
                    # 复制原始文件到 result 目录
                    shutil.copyfile(self.source_file_path, output_file_path)

                history_id = ConvertHistoryService.update_history(
                    docx_path=str(output_file_path),
                    user_ip=user_ip,
                    user_mac=user_mac,
                    success=1 if success else 0,
                    msg=msg if 'msg' in locals() else None,
                    hash_code=history_data["history_id"],
                )

                # 9️⃣ 返回结果给前端
                result_holder["result"] = {
                    "success": success,
                    "filename": output_filename,
                    "history_id": history_id,
                    "message": msg if 'msg' in locals() else None
                }

            except Exception as e:
                result_holder["result"] = {"success": False, "message": str(e)}

            finally:
                # 清理临时文件
                if self.source_file_path and os.path.exists(self.source_file_path):
                    os.unlink(self.source_file_path)
                if self.template_file_path and os.path.exists(self.template_file_path):
                    os.unlink(self.template_file_path)
                done_event.set()

        # 10️⃣ 启动线程执行转换
        thread = threading.Thread(target=run_conversion, daemon=True)
        thread.start()

        # async 等待线程完成
        while not done_event.is_set():
            await asyncio.sleep(0.1)

        return result_holder["result"]

# ============================================
# check-convert-history 接口
# ============================================
@app.post("/check-convert-history/")
async def check_convert_history(file: UploadFile = File(...)):
    """
    检查 DOCX 文件中的转换历史记录，返回规范化 JSON：
    {
        "exists": True/False,
        "history_id": str or None,
        "records": [
            {
                "hash_code": str,
                "file_path": str,
                "user_ip": str,
                "user_mac": str,
                "success": 1/0,
                "msg": str,
                "created_at": str
            },
            ...
        ]
    }
    """
    try:
        # 保存上传文件到临时文件
        suffix = Path(file.filename).suffix
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as temp_file:
            temp_file.write(await file.read())
            temp_path = temp_file.name

        # 如果是 .doc → .docx
        if Path(temp_path).suffix.lower() == ".doc":
            doc2docx.convert_doc_to_docx(temp_path)
            temp_path = str(Path(temp_path).with_suffix(".docx"))

        # 调用新 check_history，返回所有记录数组
        history_data = ConvertHistoryService.check_history(temp_path)

        # 清理临时文件
        if os.path.exists(temp_path):
            os.unlink(temp_path)

        return history_data

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/convert-document/")
async def convert_document(
        file: UploadFile = File(...),
        template_file: UploadFile | None = File(None),
        header_text: str = Form("格式化文档"),
        header_left_text: str = Form(""),
        header_right_text: str = Form(""),
        toc_title: str = Form("目录"),
        save_intermediate: bool = Form(False),
        document_type: int = Form(1),
):

    converter_service = DocumentConverterService(
        file=file,
        template_file=template_file,
        header_text=header_text,
        header_left_text=header_left_text,
        header_right_text=header_right_text,
        toc_title=toc_title,
        save_intermediate=save_intermediate,
        document_type=document_type,
    )

    try:
        await converter_service.prepare_files()
    except Exception:
        import traceback
        print("❌ prepare_files 阶段异常")
        print(traceback.format_exc())
        raise

    return StreamingResponse(
        converter_service.convert_document_stream(),
        media_type="text/event-stream",
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8002)
