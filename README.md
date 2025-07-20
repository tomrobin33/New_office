# Office-Word-MCP-Server

## 主要功能
- Word 文档内容解析（文本、表格、图片、元信息、大纲）
- Word 文档生成并上传服务器，返回公网下载链接
- **批量处理Word文档**：在内存中操作，最后统一保存，大幅提升性能，支持数万条记录、数百个表格、大量图片的大型JSON数据处理

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

### 3. auto_generate_and_upload_word（优化版本）
- 功能：**批量生成Word文档并上传**，使用内存操作，最后统一保存，大幅提升性能。**专门针对大型JSON数据处理优化，支持数万条记录、数百个表格、大量图片。**
- 参数：
  - filename: 生成的 Word 文件名
  - content: 结构化内容字典，包含标题、段落、表格、图片等
- 返回：
  - message: 生成结果描述
  - public_url: 公网下载链接
  - remote_path: 服务器路径
  - upload_result: 上传结果
  - stats: 处理统计信息（标题、段落、表格、图片数量等）
  - results: 详细处理结果

### 4. batch_generate_word_document
- 功能：**批量生成Word文档**（不上传），在内存中操作，最后统一保存，适用于大量内容处理。**专门针对大型JSON数据处理优化，支持数万条记录、数百个表格、大量图片。**
- 参数：
  - filename: 生成的 Word 文件名
  - content: 结构化内容字典
  - save_after_batch: 是否在批量处理后保存文档（默认True）
- 返回：
  - message: 生成结果描述
  - filename: 文件名
  - stats: 处理统计信息
  - results: 详细处理结果
  - saved: 是否已保存

## 依赖安装
```bash
pip install -r requirements.txt
```

## 运行方式
详见 main.py 及相关工具注册。 