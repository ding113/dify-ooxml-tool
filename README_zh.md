# Dify OOXML翻译工具

[English](README.md) | [中文](README_zh.md)

## 项目介绍

OOXML翻译工具是一个Dify插件，为Microsoft Office文档（DOCX、XLSX、PPTX）提供格式保持的翻译功能。该插件从OOXML文档中提取文本，通过LLM进行翻译，然后重建包含翻译内容的文档，同时保持所有原始格式、样式和布局。

## 核心特性

- **格式保持翻译**：保持所有原始格式、样式、图片和布局
- **多格式支持**：支持DOCX（Word）、XLSX（Excel）和PPTX（PowerPoint）文档
- **分块翻译支持**：支持单字符串和分块数组翻译工作流
- **LLM集成优化**：优化的XML格式防止翻译过程中的段落合并/分割
- **精确重建**：使用确切的XML位置跟踪进行精准文本替换
- **智能空格处理**：保持空白字符和格式细节
- **迭代节点兼容**：与Dify迭代节点无缝配合，支持并行处理

## 工具概览

本插件提供4个协同工作的翻译流水线工具：

| 工具 | 用途 | 输入 | 输出 |
|------|------|------|------|
| `extract_ooxml_text` | 从文档提取文本 | OOXML文件 + 文件ID | 文本段落及元数据 |
| `get_translation_texts` | 格式化文本供LLM使用 | 文件ID + 格式选项 | XML格式化的段落（字符串/数组） |
| `update_translations` | 更新LLM翻译结果 | 文件ID + 翻译结果 | 更新后的段落 |
| `rebuild_ooxml_document` | 生成最终文档 | 文件ID + 输入文件 | 翻译后的文档文件 |

## 工具参数

### 1. 提取OOXML文本工具

| 参数 | 类型 | 必需 | 描述 |
|------|------|------|------|
| `input_file` | file | 是 | OOXML文档文件（DOCX/XLSX/PPTX） |
| `file_id` | string | 是 | 处理会话的唯一标识符 |

### 2. 获取翻译文本工具

| 参数 | 类型 | 必需 | 默认值 | 描述 |
|------|------|------|-------|------|
| `file_id` | string | 是 | - | 来自提取步骤的文件标识符 |
| `output_format` | select | 否 | "string" | 输出格式："string"为单个文本块，"array"为分块文本 |
| `chunk_size` | number | 否 | 1500 | 每块最大字符数（当output_format="array"时） |
| `max_segments_per_chunk` | number | 否 | 50 | 每块最大XML段落数（当output_format="array"时） |

### 3. 更新翻译工具

| 参数 | 类型 | 必需 | 描述 |
|------|------|------|------|
| `file_id` | string | 是 | 文件标识符 |
| `translated_texts` | string | 否 | 来自LLM的XML格式翻译结果（支持字符串和数组格式） |

### 4. 重建OOXML文档工具

| 参数 | 类型 | 必需 | 描述 |
|------|------|------|------|
| `file_id` | string | 是 | 要重建的文档的文件标识符 |

## 输出变量

插件为工作流集成提供以下变量：

### 提取工具输出
- `success` (boolean)：提取是否成功
- `file_id` (string)：文件处理标识符
- `file_type` (string)：检测到的文档类型（docx/xlsx/pptx）
- `extracted_text_count` (integer)：找到的文本段落数量

### 获取文本工具输出
- `success` (boolean)：文本检索是否成功
- `original_texts` (string)：供LLM翻译的XML格式化段落
- `text_count` (integer)：文本段落总数

### 更新工具输出
- `success` (boolean)：翻译更新是否成功
- `updated_count` (integer)：成功更新的翻译数量
- `mismatch_warning` (boolean)：段落数量是否匹配

### 重建工具输出
- `success` (boolean)：文档重建是否成功
- `output_filename` (string)：翻译文档的生成文件名
- `download_url` (string)：用于下载的Base64编码文件内容

## 工作流集成指南

### 基础翻译工作流

创建包含以下节点序列的工作流：

```yaml
1. 文件上传节点（开始）
   ↓
2. 提取OOXML文本工具
   - input_file: {{#start.files[0]}}
   - file_id: "doc_{{#start.timestamp}}"
   ↓
3. 获取翻译文本工具
   - file_id: {{#extract.file_id}}
   ↓
4. LLM翻译节点
   - prompt: "翻译为西班牙语，保持XML格式：{{#get_texts.original_texts}}"
   ↓
5. 更新翻译工具
   - file_id: {{#extract.file_id}}
   - translated_texts: {{#llm.output}}
   ↓
6. 重建文档工具
   - file_id: {{#extract.file_id}}
   ↓
7. 输出节点（结束）
   - 下载：{{#rebuild.download_url}}
```

### LLM翻译提示词模板

在首选LLM中使用此提示词模板：

```
将以下XML段落翻译为[目标语言]。
重要：保持确切的<segment id="XXX">结构。
不要合并、分割或重新排序段落。

{{#get_texts.original_texts}}

输出格式示例：
<segment id="001">翻译文本1</segment>
<segment id="002">翻译文本2</segment>
```

### 处理多个文件

对于批量处理，使用迭代节点：

```yaml
1. 文件上传节点（多个文件）
   ↓
2. 迭代节点
   - 循环：{{#start.files}}
   - 包含：翻译工作流（上述步骤2-6）
   - 使用文件ID："batch_{{#iteration.index}}"
```


## 使用示例

### 输入文档类型
- **Word文档（.docx）**
- **Excel电子表格（.xlsx）**
- **PowerPoint演示文稿（.pptx）**

### XML格式示例

插件为LLM翻译格式化文本如下：

```xml
<segment id="001">欢迎来到我们公司</segment>
<segment id="002">2024年年度报告</segment>
<segment id="003">财务概览</segment>
```

LLM应以相同结构响应：

```xml
<segment id="001">Welcome to our company</segment>
<segment id="002">Annual Report 2024</segment>
<segment id="003">Financial Overview</segment>
```

### 输出示例

原始文档 → 翻译文档具有：
- ✅ 相同的格式和样式
- ✅ 相同的图片和图表
- ✅ 相同的布局和结构
- ✅ 翻译后的文本内容


## 贡献

欢迎贡献！请确保：
- 保持与OOXML标准的兼容性
- 包含全面的测试
- 遵循安全最佳实践
- 相应更新文档

## 许可证

本项目基于MIT许可证 - 详见[LICENSE](LICENSE)文件。