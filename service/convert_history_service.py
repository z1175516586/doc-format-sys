# service/convert_history_service.py

import zipfile
import base64
from pathlib import Path
import uuid
import xml.etree.ElementTree as ET

import mysql.connector

DB_CONFIG = {
    'host': '127.0.0.1',
    'user': 'root',
    'password': '123456',
    'database': 'szzs',
    'charset': 'utf8mb4'
}

def connect_db():
    try:
        return mysql.connector.connect(**DB_CONFIG)
    except mysql.connector.Error as err:
        print(f"[DB ERROR] {err}")
        return None

SETTINGS_XML_PATH = "word/settings.xml"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ET.register_namespace("w", W_NS)

# ======================================================
# 生成 22 字符唯一哈希
# ======================================================
def generate_short_hash() -> str:
    """
    生成 22 位 URL 安全唯一 ID
    基于 UUID4 + base64
    """
    u = uuid.uuid4().bytes
    return base64.urlsafe_b64encode(u).decode("ascii").rstrip("=")


class ConvertHistoryService:
    DOCVAR_NAME = "DocHistoryId"

    @staticmethod
    def hashcode_exists(hash_code: str) -> bool:
        """
        检查 doc_record 表中是否存在该 hash_code
        存在返回 True，不存在返回 False
        """
        import mysql.connector

        conn = None
        cursor = None
        try:
            conn = connect_db()
            if conn is None:
                return False

            cursor = conn.cursor(buffered=True)

            sql = "SELECT 1 FROM doc_record WHERE hash_code = %s LIMIT 1"
            cursor.execute(sql, (hash_code,))
            result = cursor.fetchone()

            return result is not None

        except mysql.connector.Error as err:
            print(f"[DB ERROR] hashcode_exists 查询失败: {err}")
            return False

        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    # ======================================================
    @staticmethod
    def get_file_path_by_hash(hash_code: str) -> str | None:
        """
        根据 hash_code(node_code) 查询文件路径
        """
        conn = None
        cursor = None
        try:
            conn = connect_db()
            if conn is None:
                return None

            cursor = conn.cursor()

            cursor.execute(
                "SELECT file_path FROM doc_node WHERE node_code=%s LIMIT 1",
                (hash_code,)
            )

            row = cursor.fetchone()

            if not row:
                return None

            return row[0]

        except mysql.connector.Error as err:
            print(f"[DB ERROR] get_file_path_by_hash 查询失败: {err}")
            return None

        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    # ======================================================
    @staticmethod
    def save_to_db(node_id: int,
                   user_id: int | None = None,
                   user_ip: str | None = None,
                   user_mac: str | None = None,
                   success: int = 1,
                   msg: str | None = None):

        conn = connect_db()
        if not conn:
            return

        try:
            cursor = conn.cursor()

            cursor.execute(
                """
                INSERT INTO doc_rec
                (node_id, user_id, user_ip, user_mac, success, msg)
                VALUES (%s, %s, %s, %s, %s, %s)
                """,
                (node_id, user_id, user_ip, user_mac, success, msg)
            )

            conn.commit()

        except mysql.connector.Error as e:
            print(f"[DB ERROR] {e}")

        finally:
            cursor.close()
            conn.close()

    # ======================================================
    @staticmethod
    def update_history(docx_path: str,
                       user_ip: str | None = None,
                       user_mac: str | None = None,
                       success: int = 1,
                       msg: str | None = None,
                       hash_code: str | None = None) -> str | None:
        """
        规则：
        1️⃣ 如果文档中不存在 DOCVAR_NAME → 生成新的 hash_code
        2️⃣ 如果文档中已存在 DOCVAR_NAME → 使用入参 hash_code 覆盖
        3️⃣ 每次都写入 doc_node（如不存在）和 doc_rec
        """

        suffix = Path(docx_path).suffix.lower()
        if suffix != ".docx" or not zipfile.is_zipfile(docx_path):
            print("❌ 不是合法 docx 文件")
            return None

        try:
            # 1️⃣ 读取 docx
            with zipfile.ZipFile(docx_path, "r") as zin:
                files = {item.filename: zin.read(item.filename) for item in zin.infolist()}

            if SETTINGS_XML_PATH not in files:
                print("❌ 未找到 settings.xml")
                return None

            settings_bytes = files[SETTINGS_XML_PATH]
            settings_xml = settings_bytes.decode("utf-8")

            final_code = generate_short_hash()

            new_docvar = (
                f'<w:docVar w:name="{ConvertHistoryService.DOCVAR_NAME}" '
                f'w:val="{final_code}"/>'
            )

            if "<w:docVars>" in settings_xml:
                # 已有 docVars → 插入内部
                settings_xml = settings_xml.replace(
                    "<w:docVars>",
                    f"<w:docVars>\n    {new_docvar}",
                    1
                )
            else:
                # 没有 docVars → 在 </w:compat> 后插入整个块
                insert_block = (
                    f"\n  <w:docVars>\n"
                    f"    {new_docvar}\n"
                    f"  </w:docVars>\n"
                )

                settings_xml = settings_xml.replace(
                    "</w:compat>",
                    f"</w:compat>{insert_block}",
                    1
                )

            xml_bytes = settings_xml.encode("utf-8")

            with zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for name, content in files.items():
                    if name != SETTINGS_XML_PATH:
                        zout.writestr(name, content)
                zout.writestr(SETTINGS_XML_PATH, xml_bytes)

            # ==============================
            # 6️⃣ 写入 doc_node
            # ==============================
            conn = connect_db()
            cursor = conn.cursor()

            cursor.execute("SELECT id, root_id FROM doc_node WHERE node_code=%s", (final_code,))
            row = cursor.fetchone()

            # 新节点
            parent_id = None
            root_id = None

            if hash_code:
                # 入参 hash_code 不为空 → 文档非首次 → 尝试找到父节点
                cursor.execute("SELECT id, root_id FROM doc_node WHERE node_code=%s", (hash_code,))
                parent_row = cursor.fetchone()
                if parent_row:
                    parent_id = parent_row[0]
                    root_id = parent_row[1]

            # 如果 hash_code 为空 → 首次 → root_id 之后再更新
            cursor.execute(
                """
                INSERT INTO doc_node (node_code, root_id, parent_id, file_path)
                VALUES (%s, %s, %s, %s)
                """,
                (final_code, root_id or 0, parent_id, docx_path)
            )
            node_id = cursor.lastrowid

            # 首个节点（没有父节点） → 设置 root_id 为自身
            if not parent_id:
                cursor.execute(
                    "UPDATE doc_node SET root_id=%s WHERE id=%s",
                    (node_id, node_id)
                )
                root_id = node_id

            conn.commit()
            cursor.close()
            conn.close()

            # ==============================
            # 7️⃣ 写入行为记录 doc_rec
            # ==============================
            ConvertHistoryService.save_to_db(
                node_id=node_id,
                user_id=None,
                user_ip=user_ip,
                user_mac=user_mac,
                success=success,
                msg=msg
            )

            return final_code

        except Exception as e:
            print(f"❌ update_history 失败: {e}")
            return None

    @staticmethod
    def check_history(file_path: str) -> dict:
        """
        检查 DOCX 文件的转换历史，返回：
        {
            "exists": True/False,
            "history_id": node_code,
            "records": 当前节点链路上的行为记录,
            "family_nodes": 回溯链路节点,
            "derived_nodes": 下游衍生节点,
            "is_latest": 上传文件是否为最新,
            "latest_files": 最新文件列表（可能有多个）
        }
        """

        suffix = Path(file_path).suffix.lower()
        if suffix != ".docx" or not zipfile.is_zipfile(file_path):
            return {
                "exists": False,
                "history_id": None,
                "records": [],
                "family_nodes": [],
                "derived_nodes": [],
                "is_latest": True,
                "latest_files": []
            }

        node_code = None
        try:
            with zipfile.ZipFile(file_path, "r") as z:
                xml_data = z.read(SETTINGS_XML_PATH)
                root_xml = ET.fromstring(xml_data)
                doc_vars = root_xml.find(f"{{{W_NS}}}docVars")
                if doc_vars:
                    for var in doc_vars.findall(f"{{{W_NS}}}docVar"):
                        name = var.attrib.get(f"{{{W_NS}}}name") or var.attrib.get("name")
                        val = var.attrib.get(f"{{{W_NS}}}val") or var.attrib.get("val")
                        if name == ConvertHistoryService.DOCVAR_NAME:
                            node_code = val
                            break
        except Exception:
            pass

        if not node_code:
            return {
                "exists": False,
                "history_id": None,
                "records": [],
                "family_nodes": [],
                "derived_nodes": [],
                "is_latest": True,
                "latest_files": []
            }

        conn = connect_db()
        cursor = conn.cursor(dictionary=True)

        # 获取当前节点信息
        cursor.execute("SELECT * FROM doc_node WHERE node_code=%s", (node_code,))
        node = cursor.fetchone()
        if not node:
            cursor.close()
            conn.close()
            return {
                "exists": False,
                "history_id": node_code,
                "records": [],
                "family_nodes": [],
                "derived_nodes": [],
                "is_latest": True,
                "latest_files": []
            }

        node_id = node["id"]

        # 1️⃣ 回溯链路
        chain_nodes = []
        current = node
        while current:
            chain_nodes.insert(0, current)
            if current["parent_id"]:
                cursor.execute("SELECT * FROM doc_node WHERE id=%s", (current["parent_id"],))
                current = cursor.fetchone()
            else:
                break

        # 2️⃣ 收集链路行为记录
        chain_node_ids = [n["id"] for n in chain_nodes]
        if chain_node_ids:
            format_ids = ",".join(["%s"] * len(chain_node_ids))
            sql = f"""
            SELECT r.*, n.file_path
            FROM doc_rec r
            LEFT JOIN doc_node n ON r.node_id = n.id
            WHERE r.node_id IN ({format_ids})
            ORDER BY r.created_at ASC
            """
            cursor.execute(sql, chain_node_ids)
            records = cursor.fetchall()
        else:
            records = []

        # 3️⃣ 构建链路节点信息
        family_nodes = [
            {
                "id": n["id"],
                "node_code": n["node_code"],
                "parent_id": n["parent_id"],
                "file_path": n["file_path"],
                "created_at": n["created_at"]
            }
            for n in chain_nodes
        ]

        # 4️⃣ 递归查找下游衍生节点
        def get_descendants(nid):
            cursor.execute("SELECT * FROM doc_node WHERE parent_id=%s", (nid,))
            children = cursor.fetchall()
            result = []
            for child in children:
                result.append(child)
                result.extend(get_descendants(child["id"]))
            return result

        descendants = get_descendants(node_id)

        derived_nodes = [
            {
                "id": n["id"],
                "node_code": n["node_code"],
                "parent_id": n["parent_id"],
                "file_path": n["file_path"],
                "created_at": n["created_at"]
            }
            for n in descendants
        ]

        # 5️⃣ 判断上传文件是否为最新
        latest_files = []
        is_latest = True
        if descendants:
            # 找出创建时间最大的下游节点，可能有多个同时间的最新文件
            max_created = max(n["created_at"] for n in descendants)
            latest_files = [
                {
                    "id": n["id"],
                    "node_code": n["node_code"],
                    "file_path": n["file_path"],
                    "created_at": n["created_at"]
                }
                for n in descendants if n["created_at"] == max_created
            ]
            # 上传文件不是最新
            if any(n["id"] != node_id for n in latest_files):
                is_latest = False

        cursor.close()
        conn.close()

        return {
            "exists": True,
            "history_id": node_code,
            "records": records,
            "family_nodes": family_nodes,
            "derived_nodes": derived_nodes,
            "is_latest": is_latest,
            "latest_files": latest_files
        }