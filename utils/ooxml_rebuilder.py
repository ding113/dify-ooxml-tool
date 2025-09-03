import json
import zipfile
import io
import base64
import logging
import re
from typing import Dict, List, Any, Tuple
from lxml import etree

class OOXMLRebuilder:
    """
    OOXML document rebuilder for generating translated documents.
    
    This rebuilder takes the original OOXML file and translation mappings,
    then precisely replaces text elements while preserving all formatting.
    """
    
    def __init__(self):
        """Initialize rebuilder with proper logging configuration."""
        self.logger = logging.getLogger(__name__)
        
        # Security configuration for XML parsing
        self.max_xml_size = 50 * 1024 * 1024    # 50MB per XML file limit
        
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
            'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
        }
    
    def _secure_parse_xml(self, xml_content: bytes) -> etree._Element:
        """Securely parse XML content to prevent XXE attacks."""
        if len(xml_content) > self.max_xml_size:
            raise ValueError(f"XML content too large: {len(xml_content)} bytes > {self.max_xml_size} bytes")
        
        # Use secure parser that prevents XXE attacks
        return etree.fromstring(xml_content, parser=self.xml_parser)
    
    def _should_add_space_after(self, current_text: str, next_text: str) -> bool:
        """
        判断当前文本后是否应该添加空格
        改进规则：如果当前文本和下一个文本都是类单词形式（首尾字符为字母或数字），则添加空格
        这样可以处理英文缩写、撇号等常见情况，如"don't"、"Mr."等
        
        Args:
            current_text: 当前文本
            next_text: 下一个文本
            
        Returns:
            是否需要在当前文本后添加空格
        """
        if not current_text or not next_text:
            return False
        
        current_text = current_text.strip()
        next_text = next_text.strip()
        
        if not current_text or not next_text:
            return False
        
        # 改进规则：检查首尾字符是否为字母或数字
        def is_word_like(text: str) -> bool:
            """检查文本是否为类单词形式（首尾字符为字母或数字）"""
            return text and text[0].isalnum() and text[-1].isalnum()
        
        is_current_word_like = is_word_like(current_text)
        is_next_word_like = is_word_like(next_text)
        
        return is_current_word_like and is_next_word_like
    
    def _preprocess_segments_with_spaces(self, segments: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        预处理segments，为需要添加空格的segment在translated_text末尾添加空格
        使用改进的空格检测逻辑，支持英文缩写、撇号等常见情况
        这样空格就成为segment内容的一部分，不会在后续处理中出错
        
        Args:
            segments: 所有的文本segments
            
        Returns:
            处理后的segments列表
        """
        # 按sequence_id排序确保正确的顺序
        segments.sort(key=lambda x: x.get('sequence_id', 0))
        
        # 处理每个segment，检查是否需要在其后添加空格
        for i in range(len(segments) - 1):
            current_segment = segments[i]
            next_segment = segments[i + 1]
            
            current_text = current_segment.get('translated_text', '')
            next_text = next_segment.get('translated_text', '')
            
            # 如果需要添加空格，直接修改当前segment的translated_text
            if self._should_add_space_after(current_text, next_text):
                # 确保不重复添加空格
                if current_text and not current_text.endswith(' '):
                    current_segment['translated_text'] = current_text + ' '
        
        return segments
    
    def rebuild_document(self, original_file_data: bytes, text_segments: List[Dict[str, Any]], 
                        file_type: str) -> Tuple[bytes, int]:
        """
        Rebuild OOXML document with translated text.
        
        Args:
            original_file_data: Original OOXML file bytes
            text_segments: List of text segments with translations
            file_type: "docx", "xlsx", or "pptx"
            
        Returns:
            Tuple of (new_file_bytes, replaced_count)
        """
        try:
            with zipfile.ZipFile(io.BytesIO(original_file_data), 'r') as original_zip:
                # Validate original file structure
                if not self._validate_original_structure(original_zip, file_type):
                    raise ValueError(f"Invalid {file_type} file structure")
                
                if file_type == "docx":
                    return self._rebuild_docx(original_zip, text_segments)
                elif file_type == "xlsx":
                    return self._rebuild_xlsx(original_zip, text_segments)
                elif file_type == "pptx":
                    return self._rebuild_pptx(original_zip, text_segments)
                else:
                    raise ValueError(f"Unsupported file type: {file_type}")
        except Exception as e:
            self.logger.error(f"Rebuild error for {file_type}: {str(e)}")
            raise Exception(f"Failed to rebuild {file_type} document: {str(e)}")
    
    def _rebuild_docx(self, original_zip: zipfile.ZipFile, text_segments: List[Dict[str, Any]]) -> Tuple[bytes, int]:
        """Rebuild DOCX document with translated text."""
        # Preprocess segments with simplified space insertion logic
        text_segments = self._preprocess_segments_with_spaces(text_segments)
        
        # Create new ZIP file in memory
        new_zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(new_zip_buffer, 'w', zipfile.ZIP_DEFLATED) as new_zip:
            replaced_count = 0
            
            # Group segments by XML file
            segments_by_file = self._group_segments_by_file(text_segments)
            
            # Process each file in the original ZIP
            for file_info in original_zip.filelist:
                filename = file_info.filename
                
                if filename in segments_by_file and filename.endswith('.xml'):
                    # This XML file contains text to be translated
                    xml_content = original_zip.read(filename)
                    modified_xml, file_replaced_count = self._replace_docx_text_in_xml(
                        xml_content, segments_by_file[filename]
                    )
                    new_zip.writestr(filename, modified_xml)
                    replaced_count += file_replaced_count
                else:
                    # Copy file as-is
                    new_zip.writestr(file_info, original_zip.read(filename))
        
        new_zip_buffer.seek(0)
        return new_zip_buffer.getvalue(), replaced_count
    
    def _rebuild_xlsx(self, original_zip: zipfile.ZipFile, text_segments: List[Dict[str, Any]]) -> Tuple[bytes, int]:
        """Rebuild XLSX document with translated text."""
        # Preprocess segments with simplified space insertion logic
        text_segments = self._preprocess_segments_with_spaces(text_segments)
        
        new_zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(new_zip_buffer, 'w', zipfile.ZIP_DEFLATED) as new_zip:
            replaced_count = 0
            
            # Group segments by XML file
            segments_by_file = self._group_segments_by_file(text_segments)
            
            # Process each file in the original ZIP
            for file_info in original_zip.filelist:
                filename = file_info.filename
                
                if filename == "xl/sharedStrings.xml" and filename in segments_by_file:
                    # Special handling for shared strings table
                    xml_content = original_zip.read(filename)
                    modified_xml, file_replaced_count = self._replace_xlsx_shared_strings(
                        xml_content, segments_by_file[filename]
                    )
                    new_zip.writestr(filename, modified_xml)
                    replaced_count += file_replaced_count
                elif filename in segments_by_file and filename.endswith('.xml'):
                    # Other XML files with direct text
                    xml_content = original_zip.read(filename)
                    modified_xml, file_replaced_count = self._replace_xlsx_text_in_xml(
                        xml_content, segments_by_file[filename]
                    )
                    new_zip.writestr(filename, modified_xml)
                    replaced_count += file_replaced_count
                else:
                    # Copy file as-is
                    new_zip.writestr(file_info, original_zip.read(filename))
        
        new_zip_buffer.seek(0)
        return new_zip_buffer.getvalue(), replaced_count
    
    def _rebuild_pptx(self, original_zip: zipfile.ZipFile, text_segments: List[Dict[str, Any]]) -> Tuple[bytes, int]:
        """Rebuild PPTX document with translated text."""
        # Preprocess segments with simplified space insertion logic
        text_segments = self._preprocess_segments_with_spaces(text_segments)
        
        new_zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(new_zip_buffer, 'w', zipfile.ZIP_DEFLATED) as new_zip:
            replaced_count = 0
            
            # Group segments by XML file
            segments_by_file = self._group_segments_by_file(text_segments)
            
            # Process each file in the original ZIP
            for file_info in original_zip.filelist:
                filename = file_info.filename
                
                if filename in segments_by_file and filename.endswith('.xml'):
                    # This XML file contains text to be translated
                    xml_content = original_zip.read(filename)
                    modified_xml, file_replaced_count = self._replace_pptx_text_in_xml(
                        xml_content, segments_by_file[filename]
                    )
                    new_zip.writestr(filename, modified_xml)
                    replaced_count += file_replaced_count
                else:
                    # Copy file as-is
                    new_zip.writestr(file_info, original_zip.read(filename))
        
        new_zip_buffer.seek(0)
        return new_zip_buffer.getvalue(), replaced_count
    
    def _group_segments_by_file(self, text_segments: List[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
        """Group text segments by their XML file path."""
        segments_by_file = {}
        for segment in text_segments:
            xml_path = segment.get('xml_location', {}).get('xml_file_path', '')
            # 包含所有有translated_text的segments，包括空字符串（用于替换缺失翻译）
            translated_text = segment.get('translated_text', '')
            if xml_path and translated_text is not None:
                if xml_path not in segments_by_file:
                    segments_by_file[xml_path] = []
                segments_by_file[xml_path].append(segment)
        return segments_by_file
    
    def _replace_docx_text_in_xml(self, xml_content: bytes, segments: List[Dict[str, Any]]) -> Tuple[bytes, int]:
        """Replace text in DOCX XML content using preprocessed segments with spaces."""
        try:
            root = self._secure_parse_xml(xml_content)
            replaced_count = 0
            
            # Sort segments by sequence_id to maintain order
            segments.sort(key=lambda x: x.get('sequence_id', 0))
            
            # Process segments with simplified replacement
            for segment in segments:
                # 获取完整的translated_text，包括空字符串（用于缺失翻译的替换）
                translated_text = segment.get('translated_text', '')
                if translated_text is None:  # 只跳过None值，保留空字符串用于替换原文
                    continue
                
                # Use translated_text directly - spaces already added in preprocessing
                final_text = translated_text
                
                xpath = segment.get('xml_location', {}).get('element_xpath', '')
                namespace_map = segment.get('xml_location', {}).get('namespace_map', {})
                
                if xpath:
                    try:
                        # Find the text element using xpath
                        elements = root.xpath(xpath, namespaces=namespace_map)
                        if elements:
                            elements[0].text = final_text
                            replaced_count += 1
                    except Exception as e:
                        self.logger.warning(f"Failed to replace text at xpath {xpath}: {str(e)}")
            
            return etree.tostring(root, encoding='utf-8', xml_declaration=True), replaced_count
        except Exception as e:
            self.logger.error(f"Error processing DOCX XML: {str(e)}")
            return xml_content, 0
    
    def _replace_xlsx_shared_strings(self, xml_content: bytes, segments: List[Dict[str, Any]]) -> Tuple[bytes, int]:
        """Replace text in Excel shared strings table using preprocessed segments with spaces."""
        try:
            root = self._secure_parse_xml(xml_content)
            replaced_count = 0
            
            # Sort segments by sequence_id to maintain order
            segments.sort(key=lambda x: x.get('sequence_id', 0))
            
            # Process segments with simplified replacement
            for segment in segments:
                # 获取完整的translated_text，包括空字符串（用于缺失翻译的替换）
                translated_text = segment.get('translated_text', '')
                if translated_text is None:  # 只跳过None值，保留空字符串用于替换原文
                    continue
                
                # Use translated_text directly - spaces already added in preprocessing
                final_text = translated_text
                
                xml_location = segment.get('xml_location', {})
                shared_string_index = xml_location.get('shared_string_index')
                
                if shared_string_index is not None:
                    try:
                        # FIX: 修复部分xlsx解析 - 根据保存的命名空间映射正确查找和替换元素
                        xml_location = segment.get('xml_location', {})
                        namespace_map = xml_location.get('namespace_map', {})
                        
                        # 使用正确的命名空间查找si元素
                        if 'x' in namespace_map:
                            # 使用标准Excel命名空间
                            si_elements = root.xpath(f"//x:si[{shared_string_index + 1}]", namespaces=namespace_map)
                            if si_elements:
                                si_elem = si_elements[0]
                                t_elements = si_elem.xpath(".//x:t", namespaces=namespace_map)
                                if t_elements:
                                    t_elements[0].text = final_text
                                    replaced_count += 1
                        elif 'default' in namespace_map:
                            # 使用默认命名空间
                            si_elements = root.xpath(f"//default:si[{shared_string_index + 1}]", namespaces=namespace_map)
                            if si_elements:
                                si_elem = si_elements[0]
                                t_elements = si_elem.xpath(".//default:t", namespaces=namespace_map)
                                if t_elements:
                                    t_elements[0].text = final_text
                                    replaced_count += 1
                        else:
                            # 不使用命名空间
                            si_elements = root.xpath(f"//si[{shared_string_index + 1}]")
                            if si_elements:
                                si_elem = si_elements[0]
                                t_elements = si_elem.xpath(".//t")
                                if t_elements:
                                    t_elements[0].text = final_text
                                    replaced_count += 1
                    except Exception as e:
                        self.logger.warning(f"Failed to replace shared string at index {shared_string_index}: {str(e)}")
            
            return etree.tostring(root, encoding='utf-8', xml_declaration=True), replaced_count
        except Exception as e:
            self.logger.error(f"Error processing Excel shared strings: {str(e)}")
            return xml_content, 0
    
    def _replace_xlsx_text_in_xml(self, xml_content: bytes, segments: List[Dict[str, Any]]) -> Tuple[bytes, int]:
        """Replace direct text in Excel worksheet XML using preprocessed segments with spaces."""
        try:
            root = self._secure_parse_xml(xml_content)
            replaced_count = 0
            
            # Sort segments by sequence_id to maintain order
            segments.sort(key=lambda x: x.get('sequence_id', 0))
            
            # Process segments with simplified replacement
            for segment in segments:
                # 获取完整的translated_text，包括空字符串（用于缺失翻译的替换）
                translated_text = segment.get('translated_text', '')
                if translated_text is None:  # 只跳过None值，保留空字符串用于替换原文
                    continue
                
                # Use translated_text directly - spaces already added in preprocessing
                final_text = translated_text
                
                xpath = segment.get('xml_location', {}).get('element_xpath', '')
                namespace_map = segment.get('xml_location', {}).get('namespace_map', {})
                
                if xpath:
                    try:
                        # Find the cell value element using xpath
                        elements = root.xpath(xpath, namespaces=namespace_map)
                        if elements:
                            elements[0].text = final_text
                            replaced_count += 1
                    except Exception as e:
                        self.logger.warning(f"Failed to replace text at xpath {xpath}: {str(e)}")
            
            return etree.tostring(root, encoding='utf-8', xml_declaration=True), replaced_count
        except Exception as e:
            self.logger.error(f"Error processing Excel worksheet XML: {str(e)}")
            return xml_content, 0
    
    def _replace_pptx_text_in_xml(self, xml_content: bytes, segments: List[Dict[str, Any]]) -> Tuple[bytes, int]:
        """Replace text in PowerPoint XML content using preprocessed segments with spaces."""
        try:
            root = self._secure_parse_xml(xml_content)
            replaced_count = 0
            
            # Sort segments by sequence_id to maintain order
            segments.sort(key=lambda x: x.get('sequence_id', 0))
            
            # Process segments with simplified replacement
            for segment in segments:
                # 获取完整的translated_text，包括空字符串（用于缺失翻译的替换）
                translated_text = segment.get('translated_text', '')
                if translated_text is None:  # 只跳过None值，保留空字符串用于替换原文
                    continue
                
                # Use translated_text directly - spaces already added in preprocessing
                final_text = translated_text
                
                xpath = segment.get('xml_location', {}).get('element_xpath', '')
                namespace_map = segment.get('xml_location', {}).get('namespace_map', {})
                
                if xpath:
                    try:
                        # Find the text element using xpath
                        elements = root.xpath(xpath, namespaces=namespace_map)
                        if elements:
                            elements[0].text = final_text
                            replaced_count += 1
                    except Exception as e:
                        self.logger.warning(f"Failed to replace text at xpath {xpath}: {str(e)}")
            
            return etree.tostring(root, encoding='utf-8', xml_declaration=True), replaced_count
        except Exception as e:
            self.logger.error(f"Error processing PowerPoint XML: {str(e)}")
            return xml_content, 0
    
    def _validate_original_structure(self, zip_file: zipfile.ZipFile, file_type: str) -> bool:
        """Validate the original file structure before rebuild."""
        try:
            files = zip_file.namelist()
            
            # Check for required files
            required_files = {
                "docx": ["[Content_Types].xml", "word/document.xml"],
                "xlsx": ["[Content_Types].xml", "xl/workbook.xml"],
                "pptx": ["[Content_Types].xml", "ppt/presentation.xml"]
            }
            
            if file_type in required_files:
                for required_file in required_files[file_type]:
                    if required_file not in files:
                        self.logger.error(f"Missing required file for rebuild: {required_file}")
                        return False
            
            # Test read key files to ensure they're not corrupted
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
                        # Try to parse as XML to verify integrity
                        self._secure_parse_xml(data)
                    except Exception as e:
                        self.logger.warning(f"Cannot read/parse required file {test_file}: {str(e)}")
                        return False
            
            return True
            
        except Exception as e:
            self.logger.error(f"Structure validation error: {str(e)}")
            return False