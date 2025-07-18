# Office-Word-MCP-Server

## 主要功能
- Word 文档内容解析（文本、表格、图片、元信息、大纲）
- Word 文档生成并上传服务器，返回公网下载链接

## 指令说明

### 1. process_file
- 功能：解析本地或远程 Word 文件，返回全部内容信息（文本、表格、图片、元信息、大纲），**不上传文件**。
- 参数：
  - filename: 本地文件路径（可选）
  - file_url: 远程文件 URL（可选）
  - process_type: 处理类型，默认 'extract'
- 返回：
  - text: 文档全部文本内容
  - tables: 表格内容（二维数组）
  - images: 图片信息（图片路径或描述）
  - meta: 元信息（标题、作者等）
  - outline: 结构化大纲

### 2. create_document_and_upload
- 功能：生成 Word 文件并自动上传到服务器，返回公网下载链接和服务器路径。
- 参数：
  - filename: 生成的 Word 文件名
  - title: 文档标题（可选）
  - author: 作者（可选）
- 返回：
  - message: 生成结果描述
  - public_url: 公网下载链接
  - remote_path: 服务器路径
  - upload_result: 上传结果

## 依赖安装
```bash
pip install -r requirements.txt
```

## 运行方式
详见 main.py 及相关工具注册。 