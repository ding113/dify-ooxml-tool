# Privacy Policy

## Data Processing

This OOXML Translation Tool plugin processes Microsoft Office documents (DOCX, XLSX, PPTX) for translation purposes. Here's how we handle your data:

### Document Processing

- **Document Upload**: Your uploaded documents are temporarily stored in the Dify platform's persistent storage during the translation process
- **Text Extraction**: We extract only the textual content from your documents while preserving formatting metadata
- **LLM Translation**: Extracted text is sent to your configured Large Language Model (LLM) service for translation
- **Document Reconstruction**: Translated text is used to rebuild your document with preserved formatting

### Data Storage

- **Temporary Storage**: Document data is stored temporarily using unique session identifiers (file_id)
- **Storage Duration**: Data is retained only for the duration of the translation workflow session
- **Data Removal**: No automatic cleanup is provided; storage depends on Dify platform settings
- **No Persistent Storage**: We do not permanently store your documents or translation results

### Data Sharing

- **LLM Services**: Text content is shared with your configured LLM service (OpenAI, Anthropic, etc.) for translation
- **No Third-Party Sharing**: We do not share your documents with any other third parties
- **Local Processing**: All document parsing and reconstruction happens locally within the Dify environment

### Security Measures

- **Input Validation**: Comprehensive validation prevents malicious file uploads
- **XML Security**: Protection against XXE attacks and malicious XML content
- **Memory Processing**: Documents are processed in memory without creating temporary files
- **Access Control**: Access is limited to the Dify session performing the translation

### Your Rights

- **Data Access**: You can access your data through the Dify platform
- **Data Control**: You control when and how your documents are processed
- **Session Management**: You can terminate sessions to stop processing
- **Data Portability**: Translated documents are provided for download

### Limitations

- **File Size**: Maximum file size of 100MB per document
- **File Complexity**: Maximum 10,000 internal files per OOXML document
- **Supported Formats**: Only DOCX, XLSX, and PPTX formats are supported

### Contact

For questions about this privacy policy or data handling:
- **Plugin Issues**: Open an issue in the plugin repository
- **Dify Platform**: Contact Dify support for platform-specific privacy questions
- **Data Concerns**: Report any data handling concerns through the appropriate channels

### Updates

This privacy policy may be updated to reflect changes in functionality or requirements. Users will be notified of significant changes through the plugin documentation.

---

Last Updated: January 2025