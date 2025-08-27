# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.0.4] - 2025-08-27

### 🎉 Major Features Added
- **Chunked Translation Support**: Added support for chunked translation workflows using iteration nodes
  - `get_translation_texts` now supports `output_format` parameter with "string" and "array" options
  - Array output creates optimally sized chunks (~1500 characters or 50 XML segments each)
  - `update_translations` now handles both single string and array chunk inputs
  - Full backward compatibility with existing single-string workflows

### 🔧 Technical Improvements
- **Enhanced Parameter Processing**: Advanced dynamic type detection for iteration node parameters
  - Automatic detection and parsing of serialized arrays from iteration nodes
  - Support for both real arrays and string-serialized arrays: `"['chunk1', 'chunk2']"`
  - Robust error handling with fallback to string processing
- **Storage Format Consistency**: Fixed critical storage format mismatch between tools
  - `update_translations` now preserves original storage format (batched/compressed/simple)
  - Maintains compatibility with extract tool's batched+compressed storage
  - Ensures proper data flow from extract → get_texts → update → rebuild
- **Performance Optimizations**: Enhanced JSON processing with orjson integration
  - High-performance serialization/deserialization
  - Optimized chunk combination and processing
  - Improved logging for better debugging

### 🐛 Bug Fixes
- Fixed "No translations found" error in rebuild step after successful updates
- Resolved parameter serialization issues with iteration node arrays
- Fixed storage key inconsistencies between workflow steps
- Improved error handling for malformed translation inputs

### 📚 Documentation Updates
- Added comprehensive chunked translation workflow documentation
- Updated tool descriptions with array parameter support
- Enhanced variable references for iteration node integration
- Improved error messages and troubleshooting guidance

### 🧪 Testing & Validation
- Verified end-to-end chunked translation workflow
- Tested serialized array parsing from iteration nodes
- Validated storage format preservation across all tools
- Confirmed backward compatibility with existing workflows

## [0.0.3] - 2025-08-20

### 🚀 Performance & Storage Enhancements
- Implemented gzip compression for all storage operations
- Added orjson for high-performance JSON processing
- Enhanced metadata handling and storage optimization
- Improved memory efficiency for large documents

### 🔧 Core Improvements
- Added batch processing support for large document handling
- Enhanced error handling and logging throughout the pipeline
- Improved XML parsing with better security protections
- Optimized text extraction performance

## [0.0.2] - 2025-08-20

### 🎯 Core Features
- Enhanced performance with orjson and gzip for storage optimization
- Added input_file parameter for original document in rebuild tool
- Improved metadata handling and storage efficiency

### 🔧 Technical Updates
- Optimized JSON serialization and compression
- Enhanced storage key management
- Better error handling for file operations

## [0.0.1] - 2025-01-21

### Added
- Initial release of OOXML Translation Tool for Dify
- 4-stage translation pipeline for format-preserving document translation
- Support for DOCX, XLSX, and PPTX document formats
- LLM-optimized XML formatting to prevent segment merging/splitting
- Enterprise-grade security features including ZIP bomb protection and XXE prevention
- Intelligent whitespace and formatting preservation
- Comprehensive error handling and validation
- Storage optimization to prevent stdout buffer overflow

#### Tools
- **Extract OOXML Text Tool**: Extracts translatable text while preserving XML locations
- **Get Translation Texts Tool**: Formats text segments for LLM consumption
- **Update Translations Tool**: Integrates LLM translation results with robust parsing
- **Rebuild OOXML Document Tool**: Reconstructs documents with surgical precision

#### Core Components
- **OOXML Parser Engine**: Multi-format parsing with security hardening
- **Document Rebuilder**: Format-preserving reconstruction system
- **Translation Workflow Manager**: Structured prompt engineering for LLMs
- **Persistent Storage System**: KV-based data management with optimized serialization

#### Security Features
- XML security hardening with XXE attack prevention
- ZIP bomb protection (100MB max, 10K files limit)
- Comprehensive input validation and sanitization
- Memory-controlled processing with resource limits
- No temporary file creation - in-memory processing only

#### Documentation
- Comprehensive English and Chinese README files
- Detailed tool parameter documentation
- Workflow integration guides with examples
- Privacy policy and MIT license
- Installation and troubleshooting guides

### Technical Highlights
- Namespace-aware XML processing with fallback strategies
- Precise XPath-based text replacement for format preservation
- Space information analysis for whitespace preservation
- Segment correspondence validation for translation integrity
- Performance optimizations for large document processing