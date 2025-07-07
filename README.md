# Remove '内部' from First Line Tool

本工具为Python命令行脚本，递归扫描全盘（所有磁盘）下的txt、docx、xlsx、pdf、ppt/pptx文件，将每个文件第一行中的“内部”两个字删除，并覆盖原文件。

## 使用方法
1. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```
2. 运行脚本（需管理员权限）：
   ```bash
   python main.py
   ```

## 注意事项
- 需以管理员权限运行，确保有权限访问所有磁盘。
- 建议先备份重要文件。
- 处理后文件不可恢复。
