# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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