#!/usr/bin/env python3
"""
完整的GitHub文件管理MCP服务器
"""

import asyncio
import json
import sys
import os
import base64
from typing import Any, Dict, List, Optional

# 手动读取.env文件
def load_env_file():
    """手动读取.env文件中的环境变量"""
    env_path = '.env'
    if not os.path.exists(env_path):
        print(".env文件不存在", file=sys.stderr)
        return
        
    try:
        with open(env_path, 'r', encoding='utf-8') as f:
            for line_num, line in enumerate(f, 1):
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    key, value = line.split('=', 1)
                    key = key.strip()
                    value = value.strip().strip('"\'')
                    os.environ[key] = value
                    print(f"加载环境变量: {key} = {'*' * min(len(value), 10)}...", file=sys.stderr)
    except Exception as e:
        print(f"读取.env文件时出错: {e}", file=sys.stderr)

# 加载环境变量
load_env_file()

# GitHub API 客户端类（简化版，使用requests同步调用）
class GitHubClient:
    def __init__(self, token: str):
        self.token = token
        self.base_url = "https://api.github.com"
        self.headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github.v3+json",
            "User-Agent": "MCP-GitHub-Client/1.0"
        }
    
    def search_files(self, repo: str, filename: str) -> List[Dict]:
        """搜索仓库中的文件（同步版本）"""
        try:
            import requests
            search_url = f"{self.base_url}/search/code"
            params = {
                "q": f"filename:{filename} repo:{repo}",
                "per_page": 10
            }
            
            response = requests.get(search_url, headers=self.headers, params=params)
            if response.status_code == 200:
                data = response.json()
                return data.get('items', [])
            else:
                raise Exception(f"GitHub API错误: {response.status_code} - {response.text}")
        except ImportError:
            raise Exception("需要安装requests库: pip install requests")
    
    def get_file_content(self, repo: str, path: str, ref: str = "main") -> Dict:
        """获取文件内容（同步版本）"""
        import requests
        url = f"{self.base_url}/repos/{repo}/contents/{path}"
        params = {"ref": ref}
        
        response = requests.get(url, headers=self.headers, params=params)
        if response.status_code == 200:
            return response.json()
        else:
            raise Exception(f"获取文件失败: {response.status_code} - {response.text}")

# MCP服务器框架
class MCPServer:
    def __init__(self, name: str):
        self.name = name
        self.tools = {}
        self.resources = {}
        
    def tool(self, name: str = None):
        """装饰器：注册MCP工具"""
        def decorator(func):
            tool_name = name or func.__name__
            self.tools[tool_name] = {
                'function': func,
                'name': tool_name,
                'description': func.__doc__ or '',
                'schema': self._generate_schema(func)
            }
            return func
        return decorator
    
    def _generate_schema(self, func):
        """生成工具的参数schema（简化版）"""
        return {
            "type": "object",
            "properties": {},
            "required": []
        }
    
    async def handle_request(self, request: Dict[str, Any]) -> Dict[str, Any]:
        """处理MCP请求"""
        method = request.get('method')
        params = request.get('params', {})
        request_id = request.get('id')
        
        try:
            if method == 'initialize':
                return await self._handle_initialize(params, request_id)
            elif method == 'tools/list':
                return await self._handle_list_tools(request_id)
            elif method == 'tools/call':
                return await self._handle_call_tool(params, request_id)
            else:
                return self._error_response(request_id, f"Unknown method: {method}")
        except Exception as e:
            return self._error_response(request_id, str(e))
    
    async def _handle_initialize(self, params: Dict, request_id: str):
        """处理初始化请求"""
        return {
            "jsonrpc": "2.0",
            "id": request_id,
            "result": {
                "protocolVersion": "2024-11-05",
                "capabilities": {
                    "tools": {},
                    "resources": {}
                },
                "serverInfo": {
                    "name": self.name,
                    "version": "1.0.0"
                }
            }
        }
    
    async def _handle_list_tools(self, request_id: str):
        """返回可用工具列表"""
        tools_list = []
        for tool_name, tool_info in self.tools.items():
            tools_list.append({
                "name": tool_name,
                "description": tool_info['description'],
                "inputSchema": tool_info['schema']
            })
        
        return {
            "jsonrpc": "2.0",
            "id": request_id,
            "result": {
                "tools": tools_list
            }
        }
    
    async def _handle_call_tool(self, params: Dict, request_id: str):
        """执行工具调用"""
        tool_name = params.get('name')
        arguments = params.get('arguments', {})
        
        if tool_name not in self.tools:
            return self._error_response(request_id, f"Tool not found: {tool_name}")
        
        tool_func = self.tools[tool_name]['function']
        
        try:
            # 调用工具函数
            if asyncio.iscoroutinefunction(tool_func):
                result = await tool_func(**arguments)
            else:
                result = tool_func(**arguments)
            
            return {
                "jsonrpc": "2.0",
                "id": request_id,
                "result": {
                    "content": [
                        {
                            "type": "text",
                            "text": str(result)
                        }
                    ]
                }
            }
        except Exception as e:
            return self._error_response(request_id, f"Tool execution failed: {str(e)}")
    
    def _error_response(self, request_id: str, error_message: str):
        """生成错误响应"""
        return {
            "jsonrpc": "2.0",
            "id": request_id,
            "error": {
                "code": -1,
                "message": error_message
            }
        }
    
    async def run(self):
        """启动MCP服务器，监听stdin"""
        print("MCP服务器启动中...", file=sys.stderr)
        
        while True:
            try:
                # 从stdin读取请求
                line = await asyncio.get_event_loop().run_in_executor(
                    None, sys.stdin.readline
                )
                
                if not line:
                    break
                
                # 解析JSON请求
                request = json.loads(line.strip())
                
                # 处理请求
                response = await self.handle_request(request)
                
                # 发送响应到stdout
                print(json.dumps(response), flush=True)
                
            except json.JSONDecodeError as e:
                print(f"JSON解析错误: {e}", file=sys.stderr)
            except Exception as e:
                print(f"处理请求时出错: {e}", file=sys.stderr)

# 创建GitHub文件管理服务器实例
server = MCPServer("github-file-manager")

# 从环境变量获取GitHub token
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
github_client = None

if not GITHUB_TOKEN:
    print("警告: 未设置GITHUB_TOKEN环境变量，将使用模拟数据", file=sys.stderr)
else:
    try:
        github_client = GitHubClient(GITHUB_TOKEN)
        print("GitHub客户端初始化成功", file=sys.stderr)
    except Exception as e:
        print(f"GitHub客户端初始化失败: {e}", file=sys.stderr)

def parse_csv_content(content: str, search_key: str) -> Optional[str]:
    """解析CSV内容，根据第一列查找第二列的值"""
    lines = content.strip().split('\n')
    for line in lines:
        if not line.strip():
            continue
        parts = line.split(',')
        if len(parts) >= 2:
            # 去除引号和空格
            key = parts[0].strip().strip('"\'')
            value = parts[1].strip().strip('"\'')
            if key == search_key:
                return value
    return None

# 定义工具函数
@server.tool()
def search_file_content(repo_name: str, filename: str, search_key: str):
    """
    在GitHub仓库中搜索文件内容
    
    参数:
    - repo_name: 仓库名称 (例如: "user/repo-name")
    - filename: 文件名或部分文件名
    - search_key: 要搜索的关键字（第一列的值）
    """
    if not github_client:
        return f"模拟结果: 在仓库 {repo_name} 中找到文件 {filename}，{search_key} 对应的值为: 模拟值"
    
    try:
        # 1. 搜索文件
        files = github_client.search_files(repo_name, filename)
        
        if not files:
            return f"未找到包含 '{filename}' 的文件"
        
        # 2. 遍历找到的文件，查找内容
        for file_info in files:
            file_path = file_info['path']
            try:
                # 获取文件内容
                file_data = github_client.get_file_content(repo_name, file_path)
                
                # 解码base64内容
                content = base64.b64decode(file_data['content']).decode('utf-8')
                
                # 解析CSV查找对应值
                result_value = parse_csv_content(content, search_key)
                
                if result_value:
                    return f"✅ 找到文件: {file_path}\n🔍 {search_key} 对应的值为: {result_value}"
            
            except Exception as e:
                continue  # 跳过无法处理的文件
        
        return f"❌ 在找到的文件中未发现 '{search_key}' 对应的值"
        
    except Exception as e:
        return f"❌ 搜索过程中出错: {str(e)}"

@server.tool()
def list_files(repo_name: str, path: str = ""):
    """
    列出GitHub仓库中的所有文件（测试用）
    
    参数:
    - repo_name: 仓库名称
    - path: 路径（默认为根目录）
    """
    if not github_client:
        return "模拟结果: 文件列表获取需要GitHub token"
    
    try:
        import requests
        url = f"https://api.github.com/repos/{repo_name}/contents/{path}"
        response = requests.get(url, headers=github_client.headers)
        
        if response.status_code == 200:
            files = response.json()
            file_list = []
            for file in files:
                file_list.append(f"📁 {file['name']} ({file['type']})")
            return "文件列表:\n" + "\n".join(file_list)
        else:
            return f"获取文件列表失败: {response.status_code} - {response.text}"
    except Exception as e:
        return f"错误: {str(e)}"

@server.tool()
def update_file_content(repo_name: str, filename: str, search_key: str, new_value: str):
    """
    更新GitHub仓库中的文件内容
    
    参数:
    - repo_name: 仓库名称
    - filename: 文件名
    - search_key: 要更新的行的关键字（第一列的值）
    - new_value: 新的值（第二列的值）
    """
    if not github_client:
        return f"模拟结果: 在 {repo_name}/{filename} 中将 {search_key} 的值更新为 {new_value}，PR已创建"
    
    return "更新功能开发中..."

# 如果直接运行此文件，启动MCP服务器
if __name__ == "__main__":
    asyncio.run(server.run())