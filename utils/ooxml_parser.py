import json
import zipfile
import io
import base64
import requests
import time
import logging
from typing import Dict, List, Any, Optional, Tuple
from urllib.parse import urlparse
from lxml import etree
import re

class OOXMLParser:
    """
    Core OOXML parser for extracting translatable text from DOCX, XLSX, PPTX files.
    
    This parser extracts text while preserving exact XML locations for format-preserving
    translation reconstruction.
    """
    
    def __init__(self, max_retries: int = 3, retry_delay: float = 1.0):
        """Initialize parser with error handling configuration."""
        self.max_retries = max_retries
        self.retry_delay = retry_delay
        self.logger = logging.getLogger(__name__)
        
        # Security configuration for XML parsing
        self.max_file_size = 100 * 1024 * 1024  # 100MB limit
        self.max_xml_size = 50 * 1024 * 1024    # 50MB per XML file limit
        self.max_zip_files = 10000               # Limit ZIP files to prevent ZIP bomb
        
        # Create secure XML parser that prevents XXE attacks
        self.xml_parser = etree.XMLParser(
            resolve_entities=False,  # Disable entity resolution to prevent XXE
            no_network=True,         # Disable network access
            huge_tree=False,         # Disable huge tree support
            recover=False            # Disable error recovery
        )
        
        self.namespaces = {
            # Word
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
            # Excel
            'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            # PowerPoint
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }
    
    def parse_file(self, file_data: bytes, file_type: str) -> Dict[str, Any]:
        """
        Main entry point for parsing OOXML files.
        
        Args:
            file_data: Raw file bytes
            file_type: "docx", "xlsx", or "pptx"
            
        Returns:
            Dict containing extracted text segments and metadata
        """
        # Validate inputs
        validation_result = self._validate_inputs(file_data, file_type)
        if not validation_result["valid"]:
            return {
                "success": False,
                "error": validation_result["error"],
                "text_segments": [],
                "supported_elements": []
            }
        
        # Try parsing with retry mechanism
        for attempt in range(self.max_retries + 1):
            try:
                self.logger.debug(f"Parse attempt {attempt + 1}/{self.max_retries + 1} for {file_type}")
                
                with zipfile.ZipFile(io.BytesIO(file_data), 'r') as zip_file:
                    # Validate ZIP file integrity
                    if not self._validate_zip_file(zip_file, file_type):
                        raise ValueError(f"Invalid or corrupted {file_type} file structure")
                    
                    if file_type == "docx":
                        return self._parse_docx(zip_file)
                    elif file_type == "xlsx":
                        return self._parse_xlsx(zip_file)
                    elif file_type == "pptx":
                        return self._parse_pptx(zip_file)
                    else:
                        raise ValueError(f"Unsupported file type: {file_type}")
                        
            except zipfile.BadZipFile as e:
                error_msg = f"Corrupted ZIP file (attempt {attempt + 1}): {str(e)}"
                self.logger.warning(error_msg)
                if attempt == self.max_retries:
                    return self._create_error_response(error_msg, file_type)
                    
            except (etree.XMLSyntaxError, etree.XPathEvalError) as e:
                error_msg = f"XML parsing error (attempt {attempt + 1}): {str(e)}"
                self.logger.warning(error_msg)
                if attempt == self.max_retries:
                    return self._create_error_response(error_msg, file_type)
                    
            except (MemoryError, OverflowError) as e:
                error_msg = f"Resource exhaustion (attempt {attempt + 1}): {str(e)}"
                self.logger.error(error_msg)
                return self._create_error_response(error_msg, file_type)
                
            except Exception as e:
                error_msg = f"Unexpected error (attempt {attempt + 1}): {str(e)}"
                self.logger.warning(error_msg)
                if attempt == self.max_retries:
                    return self._create_error_response(error_msg, file_type)
            
            # Wait before retry
            if attempt < self.max_retries:
                time.sleep(self.retry_delay * (attempt + 1))  # Exponential backoff
                
        return self._create_error_response("Maximum retry attempts exceeded", file_type)
    
    def _parse_docx(self, zip_file: zipfile.ZipFile) -> Dict[str, Any]:
        """Parse DOCX file and extract all text segments."""
        text_segments = []
        supported_elements = set()
        
        # Main document content - word/document.xml
        if "word/document.xml" in zip_file.namelist():
            segments = self._extract_docx_document_text(zip_file, "word/document.xml")
            text_segments.extend(segments)
            if segments:
                supported_elements.add("w:t")
        
        # Headers and footers
        for filename in zip_file.namelist():
            if filename.startswith("word/header") and filename.endswith(".xml"):
                segments = self._extract_docx_document_text(zip_file, filename)
                text_segments.extend(segments)
                if segments:
                    supported_elements.add("header_w:t")
            elif filename.startswith("word/footer") and filename.endswith(".xml"):
                segments = self._extract_docx_document_text(zip_file, filename)
                text_segments.extend(segments)
                if segments:
                    supported_elements.add("footer_w:t")
        
        # Comments
        if "word/comments.xml" in zip_file.namelist():
            segments = self._extract_docx_document_text(zip_file, "word/comments.xml")
            text_segments.extend(segments)
            if segments:
                supported_elements.add("comment_w:t")
        
        # Footnotes and endnotes
        if "word/footnotes.xml" in zip_file.namelist():
            segments = self._extract_docx_document_text(zip_file, "word/footnotes.xml")
            text_segments.extend(segments)
            if segments:
                supported_elements.add("footnote_w:t")
        
        if "word/endnotes.xml" in zip_file.namelist():
            segments = self._extract_docx_document_text(zip_file, "word/endnotes.xml")
            text_segments.extend(segments)
            if segments:
                supported_elements.add("endnote_w:t")
        
        return {
            "success": True,
            "text_segments": text_segments,
            "supported_elements": list(supported_elements),
            "file_type": "docx"
        }
    
    def _extract_docx_document_text(self, zip_file: zipfile.ZipFile, xml_path: str) -> List[Dict[str, Any]]:
        """Extract text from a specific DOCX XML file."""
        try:
            xml_content = zip_file.read(xml_path)
            root = self._secure_parse_xml(xml_content)
            text_segments = []
            
            # Find all text elements - optimize by being more specific
            text_elements = root.xpath("//w:t", namespaces=self.namespaces)
            
            for idx, text_elem in enumerate(text_elements):
                text_content = text_elem.text or ""
                # 保留所有文本元素，包括纯空格内容
                # Get parent run element for context
                run_elem = text_elem.getparent()
                paragraph_elem = run_elem.getparent() if run_elem is not None else None
                
                # Create XPath for precise location
                xpath = self._create_element_xpath(text_elem)
                
                # 分析空格信息
                space_info = self._analyze_text_spaces(text_content)
                
                text_segments.append({
                    "sequence_id": len(text_segments),
                    "text_id": f"{xml_path}_{idx}",
                    "original_text": text_content.strip(),  # 存储纯文本内容
                    "translated_text": "",
                    "xml_location": {
                        "xml_file_path": xml_path,
                        "element_xpath": xpath,
                        "parent_context": run_elem.tag if run_elem is not None else "",  # Just store tag name instead of full XML
                        "namespace_map": {"w": self.namespaces["w"]}
                    },
                    "text_metadata": {
                        "char_count": len(text_content),
                        "is_rich_text": self._has_formatting(run_elem) if run_elem is not None else False
                    },
                    "space_info": space_info
                })
            
            return text_segments
        except Exception as e:
            self.logger.error(f"Error extracting text from {xml_path}: {str(e)}")
            return []
    
    def _parse_xlsx(self, zip_file: zipfile.ZipFile) -> Dict[str, Any]:
        """Parse XLSX file and extract all text segments."""
        text_segments = []
        supported_elements = set()
        
        # Shared strings table - xl/sharedStrings.xml
        if "xl/sharedStrings.xml" in zip_file.namelist():
            segments = self._extract_xlsx_shared_strings(zip_file)
            text_segments.extend(segments)
            if segments:
                supported_elements.add("si")
        
        # Direct cell text in worksheets (rare case)
        for filename in zip_file.namelist():
            if filename.startswith("xl/worksheets/") and filename.endswith(".xml"):
                segments = self._extract_xlsx_worksheet_text(zip_file, filename)
                text_segments.extend(segments)
                if segments:
                    supported_elements.add("worksheet_text")
        
        return {
            "success": True,
            "text_segments": text_segments,
            "supported_elements": list(supported_elements),
            "file_type": "xlsx"
        }
    
    def _extract_xlsx_shared_strings(self, zip_file: zipfile.ZipFile) -> List[Dict[str, Any]]:
        """Extract text from Excel shared strings table."""
        try:
            xml_content = zip_file.read("xl/sharedStrings.xml")
            root = self._secure_parse_xml(xml_content)
            text_segments = []
            
            # FIX: 修复部分xlsx解析 - 正确处理Excel默认命名空间
            # Excel默认命名空间处理 - cache for performance
            excel_default_ns = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            
            # Try with standard Excel namespace first (more efficient)
            si_elements = root.xpath("//x:si", namespaces=excel_default_ns)
            using_excel_ns = True
            
            if not si_elements:
                # Fallback to default namespace approach
                default_ns = {'default': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                si_elements = root.xpath("//default:si", namespaces=default_ns)
                using_excel_ns = False
                if not si_elements:
                    # Final fallback - no namespace
                    si_elements = root.xpath("//si")
                    using_excel_ns = None
                    self.logger.debug(f"Excel sharedStrings fallback to no namespace - found {len(si_elements)} si elements")
                else:
                    self.logger.debug(f"Excel sharedStrings using default namespace - found {len(si_elements)} si elements")
            else:
                self.logger.debug(f"Excel sharedStrings using standard Excel namespace - found {len(si_elements)} si elements")
            
            if not si_elements:
                self.logger.warning("No shared string items found in Excel sharedStrings.xml")
                return []
            
            for si_idx, si_elem in enumerate(si_elements):
                # Use optimized namespace strategy for text elements
                if using_excel_ns is True:
                    t_elements = si_elem.xpath(".//x:t", namespaces=excel_default_ns)
                    namespace_map = excel_default_ns
                elif using_excel_ns is False:
                    t_elements = si_elem.xpath(".//default:t", namespaces=default_ns)
                    namespace_map = default_ns
                else:
                    t_elements = si_elem.xpath(".//t")
                    namespace_map = {}
                
                for t_idx, t_elem in enumerate(t_elements):
                    text_content = t_elem.text or ""
                    # 保留所有文本元素，包括纯空格内容
                    
                    # 分析空格信息
                    space_info = self._analyze_text_spaces(text_content)
                    
                    # Build XPath based on namespace strategy used
                    if using_excel_ns is True:
                        element_xpath = f"//x:si[{si_idx + 1}]//x:t[{t_idx + 1}]"
                    elif using_excel_ns is False:
                        element_xpath = f"//default:si[{si_idx + 1}]//default:t[{t_idx + 1}]"
                    else:
                        element_xpath = f"//si[{si_idx + 1}]//t[{t_idx + 1}]"
                    
                    text_segments.append({
                        "sequence_id": len(text_segments),
                        "text_id": f"shared_string_{si_idx}_{t_idx}",
                        "original_text": text_content.strip(),  # 存储纯文本内容
                        "translated_text": "",
                        "xml_location": {
                            "xml_file_path": "xl/sharedStrings.xml",
                            "element_xpath": element_xpath,
                            "shared_string_index": si_idx,  # Critical for Excel
                            "parent_context": si_elem.tag,  # Just store tag name instead of full XML
                            "namespace_map": namespace_map
                        },
                        "text_metadata": {
                            "char_count": len(text_content),
                            "is_rich_text": len(si_elem.xpath(".//r")) > 0  # Has rich text runs
                        },
                        "space_info": space_info
                    })
            
            return text_segments
        except Exception as e:
            self.logger.error(f"Error extracting shared strings: {str(e)}")
            return []
    
    def _extract_xlsx_worksheet_text(self, zip_file: zipfile.ZipFile, xml_path: str) -> List[Dict[str, Any]]:
        """Extract direct text from Excel worksheet (inline strings)."""
        try:
            xml_content = zip_file.read(xml_path)
            root = self._secure_parse_xml(xml_content)
            text_segments = []
            
            # Excel命名空间处理 - 支持多种命名空间格式
            excel_ns = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            
            # 查找内联字符串单元格 (t="str") 和其他文本单元格
            text_cells = []
            
            # Method 1: 尝试使用命名空间查询内联字符串
            inline_string_cells = root.xpath("//x:c[@t='str']/x:v", namespaces=excel_ns)
            text_cells.extend(inline_string_cells)
            
            # Method 2: 尝试使用命名空间查询内联字符串 (无引号值)
            inline_string_cells_2 = root.xpath("//x:c[@t='inlineStr']/x:is/x:t", namespaces=excel_ns)
            text_cells.extend(inline_string_cells_2)
            
            # Method 3: 如果命名空间方法失败，尝试无命名空间方法
            if not text_cells:
                # 无命名空间查询 - 内联字符串
                text_cells.extend(root.xpath("//c[@t='str']/v"))
                text_cells.extend(root.xpath("//c[@t='inlineStr']/is/t"))
                
            # Method 4: 查找所有可能包含文本的单元格(非数字，非共享字符串引用)
            if not text_cells:
                # 查找所有非共享字符串的单元格值
                all_cells = root.xpath("//x:c[not(@t='s')]/x:v", namespaces=excel_ns)
                if not all_cells:
                    all_cells = root.xpath("//c[not(@t='s')]/v")
                text_cells.extend(all_cells)
            
            self.logger.debug(f"Found {len(text_cells)} text cells in {xml_path}")
            
            for idx, cell_elem in enumerate(text_cells):
                text_content = cell_elem.text or ""
                
                # 过滤纯数字和空内容，但保留包含文字的内容
                if text_content.strip() and not text_content.replace('.', '').replace('-', '').replace('+', '').isdigit():
                    xpath = self._create_element_xpath(cell_elem)
                    
                    # 分析空格信息
                    space_info = self._analyze_text_spaces(text_content)
                    
                    text_segments.append({
                        "sequence_id": len(text_segments),
                        "text_id": f"{xml_path}_{idx}",
                        "original_text": text_content.strip(),  # 存储纯文本内容
                        "translated_text": "",
                        "xml_location": {
                            "xml_file_path": xml_path,
                            "element_xpath": xpath,
                            "parent_context": cell_elem.getparent().tag if cell_elem.getparent() is not None else "",
                            "namespace_map": {"x": self.namespaces.get("x", "")}
                        },
                        "text_metadata": {
                            "char_count": len(text_content),
                            "is_rich_text": False
                        },
                        "space_info": space_info
                    })
            
            self.logger.info(f"Extracted {len(text_segments)} text segments from {xml_path}")
            return text_segments
            
        except Exception as e:
            self.logger.error(f"Error extracting worksheet text from {xml_path}: {str(e)}")
            return []
    
    def _parse_pptx(self, zip_file: zipfile.ZipFile) -> Dict[str, Any]:
        """Parse PPTX file and extract all text segments."""
        text_segments = []
        supported_elements = set()
        
        # Slide content
        for filename in zip_file.namelist():
            if filename.startswith("ppt/slides/slide") and filename.endswith(".xml"):
                segments = self._extract_pptx_slide_text(zip_file, filename)
                text_segments.extend(segments)
                if segments:
                    supported_elements.add("slide_a:t")
        
        # Speaker notes
        for filename in zip_file.namelist():
            if filename.startswith("ppt/notesSlides/notesSlide") and filename.endswith(".xml"):
                segments = self._extract_pptx_slide_text(zip_file, filename)
                text_segments.extend(segments)
                if segments:
                    supported_elements.add("notes_a:t")
        
        # Comments
        for filename in zip_file.namelist():
            if filename.startswith("ppt/comments/comment") and filename.endswith(".xml"):
                segments = self._extract_pptx_slide_text(zip_file, filename)
                text_segments.extend(segments)
                if segments:
                    supported_elements.add("comment_a:t")
        
        return {
            "success": True,
            "text_segments": text_segments,
            "supported_elements": list(supported_elements),
            "file_type": "pptx"
        }
    
    def _extract_pptx_slide_text(self, zip_file: zipfile.ZipFile, xml_path: str) -> List[Dict[str, Any]]:
        """Extract text from PowerPoint slide XML."""
        try:
            xml_content = zip_file.read(xml_path)
            root = self._secure_parse_xml(xml_content)
            text_segments = []
            
            # Find all text elements in DrawingML - optimize query
            text_elements = root.xpath("//a:t", namespaces=self.namespaces)
            self.logger.debug(f"PPT slide {xml_path}: found {len(text_elements)} text elements")
            
            if not text_elements:
                self.logger.debug(f"No text elements found in PPT slide {xml_path}")
            
            for idx, text_elem in enumerate(text_elements):
                text_content = text_elem.text or ""
                # 保留所有文本元素，包括纯空格内容
                xpath = self._create_element_xpath(text_elem)
                
                # Get shape context - with infinite loop protection
                shape_elem = text_elem
                max_depth = 50  # Reasonable depth limit for DrawingML structure
                depth = 0
                while (shape_elem is not None and 
                       depth < max_depth and 
                       shape_elem.tag not in [
                           f"{{{self.namespaces['p']}}}sp",  # shape
                           f"{{{self.namespaces['p']}}}pic",  # picture
                           f"{{{self.namespaces['p']}}}cxnSp"  # connection shape
                       ]):
                    shape_elem = shape_elem.getparent()
                    depth += 1
                
                # Log warning if depth limit exceeded
                if depth >= max_depth:
                    self.logger.warning(f"PPT shape container search exceeded depth limit ({max_depth}) for text: {text_content[:50]}...")
                
                # 分析空格信息
                space_info = self._analyze_text_spaces(text_content)
                
                text_segments.append({
                    "sequence_id": len(text_segments),
                    "text_id": f"{xml_path}_{idx}",
                    "original_text": text_content.strip(),  # 存储纯文本内容
                    "translated_text": "",
                    "xml_location": {
                        "xml_file_path": xml_path,
                        "element_xpath": xpath,
                        "parent_context": shape_elem.tag if shape_elem is not None else "",  # Just store tag name instead of full XML
                        "namespace_map": {"a": self.namespaces["a"], "p": self.namespaces["p"]}
                    },
                    "text_metadata": {
                        "char_count": len(text_content),
                        "is_rich_text": self._has_formatting(text_elem.getparent())
                    },
                    "space_info": space_info
                })
            
            return text_segments
        except Exception as e:
            self.logger.error(f"Error extracting slide text from {xml_path}: {str(e)}")
            return []
    
    def _create_element_xpath(self, element) -> str:
        """Create an XPath expression to locate this element."""
        path_parts = []
        current = element
        
        while current is not None:
            tag = current.tag
            # Get namespace prefix
            if '}' in tag:
                namespace_uri, local_name = tag.split('}')
                namespace_uri = namespace_uri[1:]  # Remove leading '{'
                # Find prefix for this namespace
                prefix = None
                for p, uri in self.namespaces.items():
                    if uri == namespace_uri:
                        prefix = p
                        break
                if prefix:
                    tag = f"{prefix}:{local_name}"
                else:
                    tag = local_name
            
            # Get position among siblings
            siblings = [s for s in current.getparent() if s.tag == current.tag] if current.getparent() is not None else [current]
            position = siblings.index(current) + 1 if current in siblings else 1
            
            if position > 1:
                path_parts.append(f"{tag}[{position}]")
            else:
                path_parts.append(tag)
            
            current = current.getparent()
        
        path_parts.reverse()
        return "//" + "/".join(path_parts)
    
    def _has_formatting(self, element) -> bool:
        """Check if element contains formatting information."""
        if element is None:
            return False
        
        # Look for formatting elements
        formatting_tags = [
            "rPr",  # run properties
            "pPr",  # paragraph properties
            "tPr",  # text properties
            "b",    # bold
            "i",    # italic
            "u",    # underline
            "color", # color
            "sz"    # size
        ]
        
        for tag in formatting_tags:
            if element.find(f".//{tag}") is not None:
                return True
        
        return False

    @staticmethod
    def download_file(url: str) -> bytes:
        """Download file from URL."""
        try:
            response = requests.get(url, timeout=30)
            response.raise_for_status()
            return response.content
        except Exception as e:
            raise Exception(f"Failed to download file from {url}: {str(e)}")
    
    @staticmethod
    def decode_base64_file(base64_content: str) -> bytes:
        """Decode base64 file content."""
        try:
            # Remove data URL prefix if present
            if base64_content.startswith('data:'):
                base64_content = base64_content.split(',', 1)[1]
            return base64.b64decode(base64_content)
        except Exception as e:
            raise Exception(f"Failed to decode base64 content: {str(e)}")
    
    @staticmethod
    def detect_file_type(file_data: bytes) -> Optional[str]:
        """Detect OOXML file type from file content."""
        try:
            with zipfile.ZipFile(io.BytesIO(file_data), 'r') as zip_file:
                files = zip_file.namelist()
                
                if 'word/document.xml' in files:
                    return 'docx'
                elif 'xl/workbook.xml' in files:
                    return 'xlsx'
                elif 'ppt/presentation.xml' in files:
                    return 'pptx'
                else:
                    return None
        except Exception:
            return None
    
    def _validate_inputs(self, file_data: bytes, file_type: str) -> Dict[str, Any]:
        """Validate input parameters before processing."""
        if not file_data:
            return {"valid": False, "error": "Empty file data provided"}
        
        if len(file_data) < 100:  # Minimum size for a valid OOXML file
            return {"valid": False, "error": "File too small to be a valid OOXML document"}
        
        if len(file_data) > self.max_file_size:
            return {"valid": False, "error": f"File too large (>{self.max_file_size // (1024*1024)}MB)"}
        
        if file_type not in ["docx", "xlsx", "pptx"]:
            return {"valid": False, "error": f"Unsupported file type: {file_type}"}
        
        # Verify ZIP file signature
        if not file_data.startswith(b'PK'):
            return {"valid": False, "error": "Invalid file format - not a ZIP-based file"}
        
        return {"valid": True, "error": None}
    
    def _secure_parse_xml(self, xml_content: bytes) -> etree._Element:
        """Securely parse XML content to prevent XXE attacks."""
        if len(xml_content) > self.max_xml_size:
            raise ValueError(f"XML content too large: {len(xml_content)} bytes > {self.max_xml_size} bytes")
        
        # Use secure parser that prevents XXE attacks
        return etree.fromstring(xml_content, parser=self.xml_parser)
    
    def _validate_zip_file(self, zip_file: zipfile.ZipFile, file_type: str) -> bool:
        """Validate ZIP file integrity and required structure."""
        try:
            files = zip_file.namelist()
            
            # ZIP bomb protection: Check file count
            if len(files) > self.max_zip_files:
                self.logger.warning(f"ZIP file contains too many files: {len(files)} > {self.max_zip_files}")
                return False
            
            # ZIP bomb protection: Check total uncompressed size
            total_uncompressed_size = 0
            for file_info in zip_file.infolist():
                total_uncompressed_size += file_info.file_size
                if total_uncompressed_size > self.max_file_size * 10:  # Allow 10x compression ratio max
                    self.logger.warning(f"ZIP file uncompressed size too large: {total_uncompressed_size}")
                    return False
            
            # Check for required files based on file type
            required_files = {
                "docx": ["[Content_Types].xml", "word/document.xml"],
                "xlsx": ["[Content_Types].xml", "xl/workbook.xml"],
                "pptx": ["[Content_Types].xml", "ppt/presentation.xml"]
            }
            
            if file_type in required_files:
                for required_file in required_files[file_type]:
                    if required_file not in files:
                        self.logger.warning(f"Missing required file: {required_file}")
                        return False
            
            # Test read a few key files to ensure they're not corrupted
            test_files = ["[Content_Types].xml"]
            if file_type == "docx" and "word/document.xml" in files:
                test_files.append("word/document.xml")
            elif file_type == "xlsx" and "xl/workbook.xml" in files:
                test_files.append("xl/workbook.xml")
            elif file_type == "pptx" and "ppt/presentation.xml" in files:
                test_files.append("ppt/presentation.xml")
            
            for test_file in test_files:
                if test_file in files:
                    try:
                        data = zip_file.read(test_file)
                        if not data:
                            self.logger.warning(f"Empty required file: {test_file}")
                            return False
                    except Exception as e:
                        self.logger.warning(f"Cannot read required file {test_file}: {str(e)}")
                        return False
            
            return True
            
        except Exception as e:
            self.logger.error(f"ZIP validation error: {str(e)}")
            return False
    
    def _create_error_response(self, error_message: str, file_type: str) -> Dict[str, Any]:
        """Create standardized error response."""
        self.logger.error(f"Parse error for {file_type}: {error_message}")
        return {
            "success": False,
            "error": f"Failed to parse {file_type} file: {error_message}",
            "text_segments": [],
            "supported_elements": [],
            "error_context": {
                "file_type": file_type,
                "error_type": "parsing_error",
                "recovery_suggestions": [
                    "Verify file is not corrupted",
                    "Check file format is correct",
                    "Try with a smaller file for testing",
                    "Ensure file was saved properly from original application"
                ]
            }
        }
    
    def _analyze_text_spaces(self, raw_text: str) -> Dict[str, Any]:
        """
        分析文本的空格信息，用于后续空格恢复。
        
        Args:
            raw_text: 原始文本内容
            
        Returns:
            包含空格信息的字典
        """
        if not raw_text:
            return {
                "raw_text": "",
                "leading_spaces": 0,
                "trailing_spaces": 0,
                "is_pure_space": True,
                "has_content": False
            }
        
        stripped_text = raw_text.strip()
        
        return {
            "raw_text": raw_text,
            "leading_spaces": len(raw_text) - len(raw_text.lstrip()),
            "trailing_spaces": len(raw_text) - len(raw_text.rstrip()),
            "is_pure_space": stripped_text == "",
            "has_content": stripped_text != ""
        }