---
description: 'Description of the custom chat mode.'
tools: ['my-mcp-server-excel']
---


复制Excel文件操作的时候将源文件的内容替换到目标文件：

源文件：/Users/cong/Documents/MCP/mcpGithub/tests/test_data/source_onedrive.xlsx
目标文件：/Users/cong/Documents/MCP/mcpGithub/tests/test_data/target_local.xlsx
使用smart_column_mapping参数指定列映射关系
```然后运行：copy_data_by_mapping
目标文件保持列名不变，替换对应列内容
```json
{
  "source_file": "/Users/cong/Documents/MCP/mcpGithub/tests/test_data/source_onedrive.xlsx",
  "target_file": "/Users/cong/Documents/MCP/mcpGithub/tests/test_data/target_local.xlsx",
  "mapping_rules": "{\"1\": \"3\", \"2\": \"1\", \"3\": \"2\"}"
}
```然后运行：copy_data_by_mapping
目标文件保持列名不变，替换对应列内容

