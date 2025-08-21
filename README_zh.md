# Dify OOXML翻译工具

[English](README.md) | [中文](README_zh.md)

## 项目介绍

OOXML翻译工具是一个Dify插件，为Microsoft Office文档（DOCX、XLSX、PPTX）提供格式保持的翻译功能。该插件从OOXML文档中提取文本，通过LLM进行翻译，然后重建包含翻译内容的文档，同时保持所有原始格式、样式和布局。

## 核心特性

- **格式保持翻译**：保持所有原始格式、样式、图片和布局
- **多格式支持**：支持DOCX（Word）、XLSX（Excel）和PPTX（PowerPoint）文档
- **LLM集成优化**：优化的XML格式防止翻译过程中的段落合并/分割
- **企业级安全**：XML安全加固、ZIP炸弹防护和全面输入验证
- **工作流集成**：专为Dify工作流自动化设计的4阶段流水线
- **精确重建**：使用确切的XML位置跟踪进行精准文本替换
- **智能空格处理**：保持空白字符和格式细节

## 工具概览

本插件提供4个协同工作的翻译流水线工具：

| 工具 | 用途 | 输入 | 输出 |
|------|------|------|------|
| `extract_ooxml_text` | 从文档提取文本 | OOXML文件 + 文件ID | 文本段落及元数据 |
| `get_translation_texts` | 格式化文本供LLM使用 | 文件ID | XML格式化的段落 |
| `update_translations` | 更新LLM翻译结果 | 文件ID + 翻译结果 | 更新后的段落 |
| `rebuild_ooxml_document` | 生成最终文档 | 文件ID | 翻译后的文档文件 |

## 工具参数

### 1. 提取OOXML文本工具

| 参数 | 类型 | 必需 | 描述 |
|------|------|------|------|
| `input_file` | file | 是 | OOXML文档文件（DOCX/XLSX/PPTX） |
| `file_id` | string | 是 | 处理会话的唯一标识符 |

### 2. 获取翻译文本工具

| 参数 | 类型 | 必需 | 描述 |
|------|------|------|------|
| `file_id` | string | 是 | 来自提取步骤的文件标识符 |

### 3. 更新翻译工具

| 参数 | 类型 | 必需 | 描述 |
|------|------|------|------|
| `file_id` | string | 是 | 文件标识符 |
| `translated_texts` | string | 是 | 来自LLM的XML格式翻译结果 |

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

## 安装说明

### 前置要求
- Dify平台安装
- 访问用于翻译的LLM（GPT-4、Claude等）

### 安装步骤

1. **下载插件**：下载OOXML翻译工具插件包
2. **安装插件**：通过Dify插件管理界面上传
3. **创建工作流**：使用4个工具设置翻译工作流
4. **配置LLM**：连接您的首选语言模型进行翻译
5. **测试**：上传样本文档验证完整流水线

## 使用示例

### 输入文档类型
- **Word文档（.docx）**：报告、合同、文章
- **Excel电子表格（.xlsx）**：数据表、表单、报告
- **PowerPoint演示文稿（.pptx）**：幻灯片、培训材料

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

## 错误处理

插件包含针对以下情况的强大错误处理：

- **无效文件格式**：验证OOXML结构和文件完整性
- **大文件**：防止ZIP炸弹（最大100MB，10K文件）
- **翻译不匹配**：验证段落数量一致性
- **存储问题**：优雅处理持久存储失败
- **XML安全**：防止XXE攻击和恶意内容

常见错误场景和解决方案：

| 错误 | 原因 | 解决方案 |
|------|------|----------|
| "提取失败" | 无效OOXML文件 | 验证文件格式和完整性 |
| "段落数量不匹配" | LLM改变了段落结构 | 检查LLM提示词和输出格式 |
| "重建错误" | 缺少翻译数据 | 确保所有前续步骤完成 |
| "文件过大" | 文件超出限制 | 使用较小文件（<100MB，<10K组件） |

## 性能考虑

- **最佳文件大小**：50MB以下文件性能最佳
- **内存使用**：处理期间需要约3倍文件大小的内存
- **处理时间**：因文档复杂性而异（通常30秒-2分钟）
- **并发处理**：支持使用唯一file_id的多文档处理

## 支持格式

| 格式 | 扩展名 | 提取元素 |
|------|--------|----------|
| Word | .docx | 文本运行、段落、页眉、页脚 |
| Excel | .xlsx | 单元格值、共享字符串、公式 |
| PowerPoint | .pptx | 文本框、幻灯片内容、备注 |

## 安全特性

- **ZIP炸弹防护**：文件数量和大小限制
- **XML安全**：禁用外部实体解析
- **输入验证**：全面的文件格式验证
- **内存限制**：受控的资源使用
- **无临时文件**：仅内存处理

## 限制

- 最大文件大小：100MB
- 最大内部文件：每个ZIP 10,000个
- 支持格式：仅DOCX、XLSX、PPTX
- 需要LLM的结构化XML输出

## 贡献

欢迎贡献！请确保：
- 保持与OOXML标准的兼容性
- 包含全面的测试
- 遵循安全最佳实践
- 相应更新文档

## 许可证

本项目基于MIT许可证 - 详见[LICENSE](LICENSE)文件。

## 支持

- **插件问题**：在此仓库中开启issue
- **Dify集成**：查看[Dify文档](https://docs.dify.ai)
- **OOXML标准**：参考[Microsoft文档](https://docs.microsoft.com/office)

---

**作者**：ding113  
**版本**：0.0.1  
**插件类型**：工具提供者

用❤️为文档翻译工作流而制作