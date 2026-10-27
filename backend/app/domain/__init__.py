"""OOXML 领域层：基于 OOXML 标准的文档解析与改写能力。

原则：
- 全部为纯函数或轻量类，不依赖 FastAPI、不依赖数据库、不写日志到 stdout
- 所有能力都基于 DocxPackage 就地改写内存中的部件，避免反复解压/打包
- 每个能力必须可单独测试（见 backend/tests）
"""
