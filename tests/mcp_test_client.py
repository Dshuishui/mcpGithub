#!/usr/bin/env python3
"""
tests/mcp_test_client.py
MCP服务器测试客户端 - 更新版本
"""

import json
import subprocess
import sys
import time
import os

class MCPTestClient:
    def __init__(self, server_script_path):
        self.server_script_path = server_script_path
        self.server_process = None
    
    def start_server(self):
        """启动MCP服务器"""
        print("🚀 启动MCP服务器...")
        
        # 获取项目根目录（tests的上级目录）
        current_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.dirname(current_dir)
        server_path = os.path.join(project_root, self.server_script_path)
        
        if not os.path.exists(server_path):
            raise FileNotFoundError(f"服务器文件不存在: {server_path}")
        
        self.server_process = subprocess.Popen(
            [sys.executable, server_path],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            bufsize=0,
            cwd=project_root  # 设置工作目录为项目根目录
        )
        time.sleep(2)  # 等待服务器启动
        print("✅ MCP服务器已启动")
    
    def send_request(self, request):
        """发送请求到MCP服务器"""
        if not self.server_process:
            raise Exception("服务器未启动")
        
        request_json = json.dumps(request)
        print(f"📤 发送请求: {request_json}")
        
        # 发送请求
        self.server_process.stdin.write(request_json + "\n")
        self.server_process.stdin.flush()
        
        # 读取响应
        response_line = self.server_process.stdout.readline()
        if response_line:
            response = json.loads(response_line.strip())
            print(f"📥 收到响应: {json.dumps(response, indent=2, ensure_ascii=False)}")
            return response
        else:
            # 检查错误输出
            error = self.server_process.stderr.readline()
            if error:
                print(f"❌ 服务器错误: {error.strip()}")
            return None
    
    def test_initialize(self):
        """测试初始化"""
        request = {
            "jsonrpc": "2.0",
            "id": "test_init",
            "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {"name": "test-client", "version": "1.0.0"}
            }
        }
        return self.send_request(request)
    
    def test_list_tools(self):
        """测试工具列表"""
        request = {
            "jsonrpc": "2.0",
            "id": "test_list",
            "method": "tools/list",
            "params": {}
        }
        return self.send_request(request)
    
    def test_call_tool(self, tool_name, arguments):
        """测试调用工具"""
        request = {
            "jsonrpc": "2.0",
            "id": f"test_call_{tool_name}",
            "method": "tools/call",
            "params": {
                "name": tool_name,
                "arguments": arguments
            }
        }
        return self.send_request(request)
    
    def stop_server(self):
        """停止服务器"""
        if self.server_process:
            self.server_process.terminate()
            self.server_process.wait()
            print("🛑 MCP服务器已停止")

def main():
    # 创建测试客户端
    client = MCPTestClient("github_mcp_server.py")  # 服务器文件相对于项目根目录的路径
    
    try:
        # 启动服务器
        client.start_server()
        
        # 测试序列
        print("\n" + "="*50)
        print("=== 测试1: 初始化 ===")
        client.test_initialize()
        
        print("\n" + "="*50)
        print("=== 测试2: 列出工具 ===")
        response = client.test_list_tools()
        
        print("\n" + "="*50)
        print("=== 测试3: 测试GitHub工具 ===")
        client.test_call_tool("list_files", {
            "repo_name": "Dshuishui/result_nezha",
            "path": ""
        })
        
        # 测试文件路径
        source_file = "tests/test_data/source_onedrive.xlsx"
        target_file = "tests/test_data/target_local.xlsx"
        
        print("\n" + "="*50)
        print("=== 测试4: 分析源文件结构 ===")
        client.test_call_tool("analyze_table_structure", {
            "file_path": source_file
        })
        
        print("\n" + "="*50)
        print("=== 测试5: 分析目标文件结构 ===")
        client.test_call_tool("analyze_table_structure", {
            "file_path": target_file
        })
        
        print("\n" + "="*50)
        print("=== 测试6: 智能列映射分析 ===")
        client.test_call_tool("smart_column_mapping", {
            "source_file": source_file,
            "target_file": target_file
        })
        
        print("\n" + "="*50)
        print("=== 测试7: 获取列数据样本 ===")
        client.test_call_tool("get_column_data_sample", {
            "file_path": source_file,
            "column_name_or_index": "1",
            "sample_size": 3
        })

        print("\n" + "="*50)
        print("=== 测试8: 数据复制映射 ===")
        
        # 根据之前智能映射的结果定义映射规则
        # 源列1[Package Name] → 目标列3[Component Name]
        # 源列2[Component Location] → 目标列1[File Path]  
        # 源列3[DLT Category] → 目标列2[DLT Status]
        mapping_rules = {
            "1": "3",  # Package Name → Component Name
            "2": "1",  # Component Location → File Path
            "3": "2"   # DLT Category → DLT Status
        }
        
        client.test_call_tool("copy_data_by_mapping", {
            "source_file": source_file,
            "target_file": target_file,
            "mapping_rules": json.dumps(mapping_rules)
        })
        
        print("\n" + "="*50)
        print("=== 测试9: 验证复制结果 ===")
        print("重新分析目标文件，查看复制后的数据：")
        client.test_call_tool("analyze_table_structure", {
            "file_path": target_file
        })

        # 在 tests/mcp_test_client.py 的 main() 函数最后添加

        print("\n" + "="*50)
        print("=== 测试10: 对比Excel文件差异 ===")
        print("对比源文件(OneDrive)和目标文件(本地扫描结果)的差异:")
        
        # 注意：这里我们对比的是原始文件，而不是复制后的文件
        # 因为我们想看到真实的差异分析
        original_target = "tests/test_data/target_local.xlsx"
        
        # 为了更好的测试效果，我们使用第1列作为关键列进行对比
        # 对于源文件，第1列是 Package Name
        # 但我们要基于文件路径进行对比，所以使用第2列 (Component Location)
        client.test_call_tool("compare_excel_files", {
            "file1": source_file,      # OneDrive文件
            "file2": original_target,  # 原始的本地文件
            "key_column": "2"          # 使用第2列(文件路径)作为对比关键字
        })
        
    except KeyboardInterrupt:
        print("\n⚠️  用户中断测试")
    except Exception as e:
        print(f"❌ 测试过程出错: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # 停止服务器
        client.stop_server()

if __name__ == "__main__":
    main()