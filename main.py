import os
import sys
import logging
from datetime import datetime

# ==============================
# 路径设置
# ==============================
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

from service.converter import DocumentConverter


def setup_logging():
    """配置日志并重定向 stdout / stderr"""
    log_dir = r"C:\DevOps\doc_convertor_wjx\logs"
    os.makedirs(log_dir, exist_ok=True)

    time_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(log_dir, f"convert_{time_str}.log")

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.__stdout__)  # 可选：同时输出到控制台
        ]
    )

    # 将 print 重定向到 logging
    class StreamToLogger:
        def __init__(self, logger, level):
            self.logger = logger
            self.level = level

        def write(self, message):
            if message.strip():
                self.logger.log(self.level, message.strip())

        def flush(self):
            pass

    sys.stdout = StreamToLogger(logging.getLogger("STDOUT"), logging.INFO)
    sys.stderr = StreamToLogger(logging.getLogger("STDERR"), logging.ERROR)

    logging.info("日志系统初始化完成")
    logging.info(f"日志文件路径：{log_file}")


def main():
    setup_logging()

    logging.info("开始文档转换任务")

    source_file = (
        r"C:\DevOps\doc_convertor_wjx\init_docx"
        r"\二烧翻车机可研汇总20240930.docx"
    )

    result_dir = os.path.join(current_dir, "result")
    os.makedirs(result_dir, exist_ok=True)

    time_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(result_dir, f"formatted_document_{time_str}.docx")

    logging.info(f"源文件：{source_file}")
    logging.info(f"输出文件：{output_file}")

    with DocumentConverter() as converter:
        success = converter.convert_document(
            source_file=source_file,
            output_file=output_file,
            header_text="数字总师可行性报告",
            toc_title="目 录",
            save_intermediate=True
        )

    if success:
        logging.info("文档转换成功")
    else:
        logging.error("文档转换失败")


if __name__ == "__main__":
    main()

# 先别删除临时目录 看看 页脚 如何 处理的