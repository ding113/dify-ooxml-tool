# OOXML Translation Tool for Dify

[English](README.md) | [中文](README_zh.md)

## Description

OOXML Translation Tool is a Dify plugin that provides format-preserving translation for Microsoft Office documents (DOCX, XLSX, PPTX). This plugin extracts text from OOXML documents, facilitates LLM-based translation, and reconstructs the document with translated content while maintaining all original formatting, styles, and layout.

## Features

- **Format-Preserving Translation**: Maintains all original formatting, styles, images, and layout
- **Multi-Format Support**: Supports DOCX (Word), XLSX (Excel), and PPTX (PowerPoint) documents
- **LLM Integration**: Optimized XML format prevents segment merging/splitting during translation
- **Precise Reconstruction**: Uses exact XML location tracking for surgical text replacement
- **Intelligent Space Handling**: Preserves whitespace and formatting nuances

## Tools Overview

This plugin provides 4 tools that work together in a translation pipeline:

| Tool | Purpose | Input | Output |
|------|---------|-------|--------|
| `extract_ooxml_text` | Extract text from document | OOXML file + file_id | Text segments with metadata |
| `get_translation_texts` | Format text for LLM | file_id | XML-formatted segments |
| `update_translations` | Update with LLM results | file_id + translations | Updated segments |
| `rebuild_ooxml_document` | Generate final document | file_id | Translated document file |

## Tool Parameters

### 1. Extract OOXML Text Tool

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `input_file` | file | Yes | OOXML document file (DOCX/XLSX/PPTX) |
| `file_id` | string | Yes | Unique identifier for processing session |

### 2. Get Translation Texts Tool

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `file_id` | string | Yes | File identifier from extraction step |

### 3. Update Translations Tool

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `file_id` | string | Yes | File identifier |
| `translated_texts` | string | Yes | XML-formatted translations from LLM |

### 4. Rebuild OOXML Document Tool

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `file_id` | string | Yes | File identifier for document to rebuild |

## Output Variables

The plugin provides these variables for workflow integration:

### Extract Tool Output
- `success` (boolean): Whether extraction was successful
- `file_id` (string): File processing identifier  
- `file_type` (string): Detected document type (docx/xlsx/pptx)
- `extracted_text_count` (integer): Number of text segments found

### Get Texts Tool Output  
- `success` (boolean): Whether text retrieval was successful
- `original_texts` (string): XML-formatted segments for LLM translation
- `text_count` (integer): Total number of text segments

### Update Tool Output
- `success` (boolean): Whether translation update was successful
- `updated_count` (integer): Number of translations successfully updated
- `mismatch_warning` (boolean): Whether segment counts matched

### Rebuild Tool Output
- `success` (boolean): Whether document rebuild was successful
- `output_filename` (string): Generated filename for translated document
- `download_url` (string): Base64-encoded file content for download

## Workflow Integration Guide

### Basic Translation Workflow

Create a workflow with the following node sequence:

```yaml
1. File Upload Node (Start)
   ↓
2. Extract OOXML Text Tool
   - input_file: {{#start.files[0]}}
   - file_id: "doc_{{#start.timestamp}}"
   ↓
3. Get Translation Texts Tool  
   - file_id: {{#extract.file_id}}
   ↓
4. LLM Translation Node
   - prompt: "Translate to Spanish, maintain XML format: {{#get_texts.original_texts}}"
   ↓
5. Update Translations Tool
   - file_id: {{#extract.file_id}}
   - translated_texts: {{#llm.output}}
   ↓
6. Rebuild Document Tool
   - file_id: {{#extract.file_id}}
   ↓
7. Output Node (End)
   - Download: {{#rebuild.download_url}}
```

### LLM Translation Prompt Template

Use this prompt template with your preferred LLM:

```
Translate the following XML segments to [TARGET_LANGUAGE].
IMPORTANT: Maintain the exact <segment id="XXX"> structure.
Do not merge, split, or reorder segments.

{{#get_texts.original_texts}}

Output format example:
<segment id="001">Translated text 1</segment>
<segment id="002">Translated text 2</segment>
```

### Processing Multiple Files

For batch processing, use an iteration node:

```yaml
1. File Upload Node (Multiple files)
   ↓
2. Iteration Node
   - Loop through: {{#start.files}}
   - Contains: Translation workflow (steps 2-6 above)
   - Use file_id: "batch_{{#iteration.index}}"
```

## Usage Examples

### Input Document Types
- **Word Documents (.docx)**
- **Excel Spreadsheets (.xlsx)**
- **PowerPoint Presentations (.pptx)**

### XML Format Example

The plugin formats text for LLM translation like this:

```xml
<segment id="001">Welcome to our company</segment>
<segment id="002">Annual Report 2024</segment>
<segment id="003">Financial Overview</segment>
```

LLM should respond with the same structure:

```xml
<segment id="001">Bienvenido a nuestra empresa</segment>
<segment id="002">Informe Anual 2024</segment>
<segment id="003">Resumen Financiero</segment>
```

### Output Example

Original Document → Translated Document with:
- ✅ Same formatting and styles
- ✅ Same images and charts
- ✅ Same layout and structure
- ✅ Translated text content

## Contributing

Contributions welcome! Please ensure:
- Maintain compatibility with OOXML standards
- Include comprehensive tests
- Follow security best practices
- Update documentation accordingly

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.