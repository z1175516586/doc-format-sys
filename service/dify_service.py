"""
Dify服务封装类
提供文件上传和工作流执行功能
"""

import requests
import os
from typing import Optional, Dict, Any
import json
import re


class DifyService:
    """Dify服务封装类"""

    def __init__(self, api_key: str, base_url: str = "http://localhost/v1"):
        """
        初始化Dify服务

        Args:
            api_key (str): Dify API密钥
            base_url (str): Dify服务基础URL
        """
        self.api_key = api_key
        self.base_url = base_url.rstrip('/')  # 移除末尾的斜杠
        self.headers = {
            "Authorization": f"Bearer {api_key}"
        }
        
        # 验证API密钥格式
        self._validate_api_key()

    def _validate_api_key(self) -> None:
        """验证API密钥格式"""
        if not self.api_key:
            raise ValueError("API密钥不能为空")
        
        # # 检查是否为标准格式(app-前缀+26位字符)
        # if not self.api_key.startswith("app-") or len(self.api_key) != 30:
        #     print(f"警告: API密钥格式可能不正确，应为30位字符且以'app-'开头，当前长度: {len(self.api_key)}")

    def upload_file(self, file_path: str, user: str = "difyuser") -> Optional[str]:
        """
        上传文件到Dify服务

        Args:
            file_path (str): 本地文件路径
            user (str): 用户标识

        Returns:
            Optional[str]: 文件ID，上传失败返回None
        """
        if not os.path.exists(file_path):
            print(f"错误: 文件不存在: {file_path}")
            return None
            
        if not os.path.isfile(file_path):
            print(f"错误: 路径不是文件: {file_path}")
            return None

        upload_url = f"{self.base_url}/files/upload"
        
        # 更新headers，移除Content-Type以便requests自动设置
        headers = self.headers.copy()

        try:
            print(f"上传文件中: {file_path}")
            filename = os.path.basename(file_path)
            with open(file_path, 'rb') as f:
                files = {
                    'file': (
                        filename,
                        f,
                        'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    )
                }
                data = {
                    "user": user,
                    "type": "document"
                }

                resp = requests.post(upload_url, headers=headers, files=files, data=data, timeout=30)
                if resp.status_code == 201:
                    print("文件上传成功")
                    return resp.json().get("id")
                else:
                    print(f"文件上传失败，状态码: {resp.status_code}")
                    # 打印服务端明确的错误提示
                    try:
                        error_info = resp.json()
                        print("错误详情：", json.dumps(error_info, ensure_ascii=False, indent=2))
                    except Exception:
                        print("错误详情：", resp.text)
                    return None
        except FileNotFoundError:
            print(f"文件未找到: {file_path}")
            return None
        except requests.exceptions.Timeout:
            print("请求超时")
            return None
        except requests.exceptions.ConnectionError:
            print("连接错误，请检查Dify服务是否运行")
            return None
        except Exception as e:
            print(f"发生错误: {str(e)}")
            return None

    def run_workflow(self, file_id: str, user: str = "difyuser", 
                     response_mode: str = "blocking") -> Dict[str, Any]:
        """
        运行Dify工作流

        Args:
            file_id (str): 已上传文件的ID
            user (str): 用户标识
            response_mode (str): 响应模式

        Returns:
            Dict[str, Any]: 工作流执行结果
        """
        workflow_url = f"{self.base_url}/workflows/run"
        headers = self.headers.copy()
        headers["Content-Type"] = "application/json"

        data = {
            "inputs": {
                "file": {
                    "transfer_method": "local_file",
                    "upload_file_id": file_id,
                    "type": "document"
                }
            },
            "response_mode": response_mode,
            "user": user
        }

        try:
            print("运行工作流...")
            response = requests.post(workflow_url, headers=headers, json=data, timeout=60)
            if response.status_code == 200:
                print("工作流执行成功")
                return response.json()
            else:
                print(f"工作流执行失败，状态码: {response.status_code}")
                try:
                    error_info = response.json()
                    print("错误详情：", json.dumps(error_info, ensure_ascii=False, indent=2))
                except Exception:
                    print("错误详情：", response.text)
                return {"status": "error", "message": f"Failed to execute workflow, status code: {response.status_code}"}
        except requests.exceptions.Timeout:
            print("请求超时")
            return {"status": "error", "message": "Request timeout"}
        except requests.exceptions.ConnectionError:
            print("连接错误，请检查Dify服务是否运行")
            return {"status": "error", "message": "Connection error"}
        except Exception as e:
            print(f"发生错误: {str(e)}")
            return {"status": "error", "message": str(e)}

    def process_document(self, file_path: str, user: str = "difyuser") -> str:
        """
        处理文档：上传文件并运行工作流，返回JSON格式的结果

        Args:
            file_path (str): 本地文件路径
            user (str): 用户标识

        Returns:
            str: JSON格式的处理结果
        """
        # 上传文件
        file_id = self.upload_file(file_path, user)
        if not file_id:
            result = {"status": "error", "message": "文件上传失败"}
            return json.dumps(result, ensure_ascii=False, indent=2)

        # 运行工作流
        result = self.run_workflow(file_id, user)
        return json.dumps(result, ensure_ascii=False, indent=2)

    def extract_outputs_content(self, file_path: str, user: str = "difyuser") -> str:
        """
        处理文档并提取outputs中的内容，清理特殊字符和大模型思考内容

        Args:
            file_path (str): 本地文件路径
            user (str): 用户标识

        Returns:
            str: 清理后的outputs内容
        """
        # 获取完整的处理结果
        result_json = self.process_document(file_path, user)
        
        # 初始化outputs_text变量
        outputs_text = ""
        
        try:
            # 解析JSON
            result = json.loads(result_json)
            
            # 检查是否有错误
            if result.get("status") == "error":
                return f"处理文档时发生错误: {result.get('message', '未知错误')}"
            
            # 提取outputs中的text内容
            outputs_text = result.get("data", {}).get("outputs", {}).get("text", "")
            
            # 清理内容：移除开始和结尾的换行符
            cleaned_text = outputs_text.strip()
            
            # 移除字符串中的控制字符，防止JSON解析错误
            # 只保留可打印字符和常见的空白字符
            cleaned_chars = []
            for char in cleaned_text:
                # 保留常见的可打印字符和空白字符
                if (32 <= ord(char) <= 126) or (char in '\n\r\t') or (ord(char) >= 128):
                    cleaned_chars.append(char)
                # 特别处理一些常见的控制字符
                elif char == '\u0000':  # 空字符
                    continue
                elif char == '\u0001':  # 标题开始
                    continue
                elif char == '\u0002':  # 文本结束
                    continue
                elif char == '\u0003':  # 传输结束
                    continue
                elif char == '\u0004':  # 等待
                    continue
                elif char == '\u0005':  # 请求发送
                    continue
                elif char == '\u0006':  # 确认
                    continue
                elif char == '\u0007':  # 响铃
                    continue
                elif char == '\u0008':  # 退格
                    continue
                elif char == '\u000b':  # 垂直制表符
                    continue
                elif char == '\u000c':  # 换页
                    continue
                elif char == '\u000e':  # 转换
                    continue
                elif char == '\u000f':  # 复位
                    continue
                elif char == '\u0010':  # 数据链路转义
                    continue
                elif char == '\u0011':  # 设备控制1
                    continue
                elif char == '\u0012':  # 设备控制2
                    continue
                elif char == '\u0013':  # 设备控制3
                    continue
                elif char == '\u0014':  # 设备控制4
                    continue
                elif char == '\u0015':  # 否定确认
                    continue
                elif char == '\u0016':  # 同步空闲
                    continue
                elif char == '\u0017':  # 传输块结束
                    continue
                elif char == '\u0018':  # 取消
                    continue
                elif char == '\u0019':  # 介质结束
                    continue
                elif char == '\u001a':  # 替换
                    continue
                elif char == '\u001b':  # 转义
                    continue
                elif char == '\u001c':  # 文件分隔符
                    continue
                elif char == '\u001d':  # 组分隔符
                    continue
                elif char == '\u001e':  # 记录分隔符
                    continue
                elif char == '\u001f':  # 单元分隔符
                    continue
            
            cleaned_text = ''.join(cleaned_chars)
            
            # 删除大模型思考内容（在第一个[之前的所有内容）
            first_bracket_index = cleaned_text.find('[')
            if first_bracket_index > 0:
                # 保留第一个[字符，只删除之前的内容
                cleaned_text = cleaned_text[first_bracket_index:]
            
            # 删除大模型思考内容（在最后一个]之后的所有内容）
            last_bracket_index = cleaned_text.rfind(']')
            if last_bracket_index != -1 and last_bracket_index < len(cleaned_text) - 1:
                # 保留最后一个]字符，只删除之后的内容
                cleaned_text = cleaned_text[:last_bracket_index+1]
            
            return cleaned_text
            
        except json.JSONDecodeError as e:
            # 如果JSON解析失败，返回原始内容以便调试
            print(f"JSON解析错误详情: {str(e)}")
            # 返回清理前的内容用于调试
            return f"JSON解析错误: {str(e)}, 原始内容: {outputs_text[:200]}..."
        except Exception as e:
            return f"处理过程中发生错误: {str(e)}"

    def extract_cover(self, file_path: str, user: str = "difyuser") -> str:
        """
        提取文档封面信息，使用upload_file和run_workflow，然后对结果进行处理，去掉大模型思考的部分

        Args:
            file_path (str): 本地文件路径
            user (str): 用户标识

        Returns:
            str: JSON格式的封面信息
        """
        # 上传文件
        file_id = self.upload_file(file_path, user)
        if not file_id:
            result = {"status": "error", "message": "文件上传失败"}
            return json.dumps(result, ensure_ascii=False, indent=2)

        # 运行工作流
        result = self.run_workflow(file_id, user)
        
        try:
            # 检查是否有错误
            if result.get("status") == "error":
                return json.dumps(result, ensure_ascii=False, indent=2)
            
            # 提取outputs中的text内容
            outputs_text = result.get("data", {}).get("outputs", {}).get("text", "")
            
            # 清理内容：移除开始和结尾的换行符
            cleaned_text = outputs_text.strip()
            
            # 删除大模型思考内容（在第一个{之前的所有内容）
            first_brace_index = cleaned_text.find('{')
            if first_brace_index > 0:
                # 保留第一个{字符，只删除之前的内容
                cleaned_text = cleaned_text[first_brace_index:]
            
            # 删除大模型思考内容（在最后一个}之后的所有内容）
            last_brace_index = cleaned_text.rfind('}')
            if last_brace_index != -1 and last_brace_index < len(cleaned_text) - 1:
                # 保留最后一个}字符，只删除之后的内容
                cleaned_text = cleaned_text[:last_brace_index+1]
            
            # 验证是否为有效的JSON
            try:
                cover_info = json.loads(cleaned_text)
                return json.dumps(cover_info, ensure_ascii=False, indent=2)
            except json.JSONDecodeError:
                # 如果不是有效的JSON，返回清理后的文本
                return cleaned_text
                
        except Exception as e:
            error_result = {"status": "error", "message": f"处理过程中发生错误: {str(e)}"}
            return json.dumps(error_result, ensure_ascii=False, indent=2)

    def get_file_info(self, file_id: str) -> Dict[str, Any]:
        """
        获取文件信息

        Args:
            file_id (str): 文件ID

        Returns:
            Dict[str, Any]: 文件信息
        """
        file_url = f"{self.base_url}/files/{file_id}"
        headers = self.headers.copy()
        headers["Content-Type"] = "application/json"

        try:
            response = requests.get(file_url, headers=headers, timeout=30)
            if response.status_code == 200:
                return response.json()
            else:
                print(f"获取文件信息失败，状态码: {response.status_code}")
                return {"status": "error", "message": f"Failed to get file info, status code: {response.status_code}"}
        except Exception as e:
            print(f"发生错误: {str(e)}")
            return {"status": "error", "message": str(e)}


# 使用示例
if __name__ == "__main__":
    # 初始化Dify服务
    try:
        dify_service = DifyService(api_key="app-x6Yx0iAI3aUFMpGS9fpJQowq", base_url="http://172.23.109.63:8088/v1")

        # 测试文件路径
        file_path = r"C:\Users\MI\AppData\Local\Temp\doc_converter_8qhapvjs\tmp1bfzdulw_预处理标题_封面.docx"

        # 处理文档
        if os.path.exists(file_path):
            print("开始处理文档...")
            result_json = dify_service.process_document(file_path)
            print("处理结果:")
            print(result_json)

            # 提取outputs内容
            print("\n提取的outputs内容:")
            outputs_content = dify_service.extract_outputs_content(file_path)
            print(outputs_content)
        else:
            print(f"测试文件不存在: {file_path}")

    except Exception as e:
        print(f"初始化Dify服务失败: {str(e)}")