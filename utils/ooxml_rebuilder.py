import json
import zipfile
import io
import base64
import logging
import re
from typing import Dict, List, Any, Tuple
from lxml import etree

# Configure logger for debug output - enable debug logs for space processing
# Note: To see debug logs, set LOG_LEVEL=DEBUG in environment or configure the logger accordingly
ooxml_rebuilder_logger = logging.getLogger(__name__)
# Set to DEBUG level to enable detailed space processing logs
ooxml_rebuilder_logger.setLevel(logging.DEBUG)

class OOXMLRebuilder:
    """
    OOXML document rebuilder for generating translated documents.
    
    This rebuilder takes the original OOXML file and translation mappings,
    then precisely replaces text elements while preserving all formatting.
    """
    
    def __init__(self):
        """Initialize rebuilder with proper logging configuration."""
        # Use the pre-configured logger with DEBUG level for space processing
        self.logger = ooxml_rebuilder_logger
        
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
        规则：只检查连接点字符 - 当前文本的尾字符是字母或数字，且下一个文本的首字符是字母或数字时添加空格
        这样可以更精确地处理各种情况，避免在标点符号连接处添加不必要的空格
        
        Args:
            current_text: 当前文本
            next_text: 下一个文本
            
        Returns:
            是否需要在当前文本后添加空格
        """
        # Debug: 记录输入参数
        self.logger.debug(f"[OOXMLRebuilder] Space check - Current: {repr(current_text)}, Next: {repr(next_text)}")
        
        if not current_text or not next_text:
            self.logger.debug(f"[OOXMLRebuilder] Space check - Empty text detected, no space needed")
            return False
        
        current_text_stripped = current_text.strip()
        next_text_stripped = next_text.strip()
        
        # Debug: 记录strip后的结果
        self.logger.debug(f"[OOXMLRebuilder] Space check - After strip - Current: {repr(current_text_stripped)}, Next: {repr(next_text_stripped)}")
        
        if not current_text_stripped or not next_text_stripped:
            self.logger.debug(f"[OOXMLRebuilder] Space check - Empty text after strip, no space needed")
            return False
        
        # 提取连接点字符
        current_last_char = current_text_stripped[-1]
        next_first_char = next_text_stripped[0]
        
        # 判断字符类型
        current_is_alnum = current_last_char.isalnum()
        next_is_alnum = next_first_char.isalnum()
        
        # Debug: 记录字符分析
        self.logger.debug(f"[OOXMLRebuilder] Space check - Connection chars: '{current_last_char}' (alnum: {current_is_alnum}) -> '{next_first_char}' (alnum: {next_is_alnum})")
        
        # 只检查连接点字符：前文本的尾字符和后文本的首字符
        needs_space = current_is_alnum and next_is_alnum
        
        # Debug: 记录最终决策
        self.logger.debug(f"[OOXMLRebuilder] Space check - Decision: {'ADD SPACE' if needs_space else 'NO SPACE'} (rule: alnum->alnum = {needs_space})")
        
        return needs_space

    def _should_add_space_after_punct(self, current_text: str, next_text: str) -> bool:
        """扩展规则：处理英文标点/引号/括号的常见空格需求。

        - 标点（, . ; : ! ? 及右括号/引号）后接字母/数字：需要空格
        - 字母/数字 后接 开引号/左括号：需要空格
        仅用于补充 `_should_add_space_after` 未覆盖的情形。
        """
        try:
            if not current_text or not next_text:
                return False
            ct = current_text.strip()
            nt = next_text.strip()
            if not ct or not nt:
                return False

            cl = ct[-1]
            nf = nt[0]
            next_is_alnum = nf.isalnum()

            punctuation_trailing = set(",.;:!?)]}\"'")
            if cl in punctuation_trailing and next_is_alnum:
                return True

            opening_punct = set("([{\"'")
            if ct[-1].isalnum() and nf in opening_punct:
                return True

            return False
        except Exception:
            return False

    def _requires_xml_space_preserve(self, text: str) -> bool:
        """判断是否需要通过 xml:space="preserve" 保留空格。"""
        if text is None or not isinstance(text, str):
            return False
        return text.startswith(' ') or text.endswith(' ') or ('  ' in text)

    def _apply_xml_space_preserve(self, element, text: str) -> None:
        """在元素上按需设置 xml:space="preserve" 以保留空格。"""
        try:
            ns_attr = '{http://www.w3.org/XML/1998/namespace}space'
            if self._requires_xml_space_preserve(text):
                element.set(ns_attr, 'preserve')
                self.logger.debug("[OOXMLRebuilder] xml:space='preserve' applied to element")
            else:
                if ns_attr in element.attrib:
                    del element.attrib[ns_attr]
        except Exception as e:
            self.logger.debug(f"[OOXMLRebuilder] Failed to set xml:space preserve: {e}")
    
    def _preprocess_segments_with_spaces(self, segments: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        预处理segments，为需要添加空格的segment在translated_text末尾添加空格
        使用改进的空格检测逻辑，只检查连接点字符，更精确地处理各种情况
        这样空格就成为segment内容的一部分，不会在后续处理中出错
        
        Args:
            segments: 所有的文本segments
            
        Returns:
            处理后的segments列表
        """
        # Debug: 记录开始处理
        self.logger.debug(f"[OOXMLRebuilder] Starting space preprocessing for {len(segments)} segments")
        
        # 按sequence_id排序确保正确的顺序
        segments.sort(key=lambda x: x.get('sequence_id', 0))
        
        # Debug: 记录排序后的segments信息
        if segments:
            first_seq = segments[0].get('sequence_id', 'N/A')
            last_seq = segments[-1].get('sequence_id', 'N/A')
            self.logger.debug(f"[OOXMLRebuilder] Segments sorted by sequence_id: {first_seq} -> {last_seq}")
        
        spaces_added_count = 0
        processed_pairs_count = 0
        
        # 处理每个segment，检查是否需要在其后添加空格
        for i in range(len(segments) - 1):
            current_segment = segments[i]
            next_segment = segments[i + 1]
            
            current_seq_id = current_segment.get('sequence_id', 'N/A')
            next_seq_id = next_segment.get('sequence_id', 'N/A')
            current_text = current_segment.get('translated_text', '')
            next_text = next_segment.get('translated_text', '')
            
            processed_pairs_count += 1
            
            # Debug: 记录每个segment的处理情况
            self.logger.debug(f"[OOXMLRebuilder] Processing segment pair {i+1}/{len(segments)-1}: seq_{current_seq_id} -> seq_{next_seq_id}")
            self.logger.debug(f"[OOXMLRebuilder] Segment texts: {repr(current_text)} -> {repr(next_text)}")
            
            # 如果需要添加空格，直接修改当前segment的translated_text
            should_add = (
                self._should_add_space_after(current_text, next_text) or
                self._should_add_space_after_punct(current_text, next_text)
            )
            # 特例：数字与数字相邻时不加空格
            if should_add:
                ct_s = (current_text or '').strip()
                nt_s = (next_text or '').strip()
                if ct_s and nt_s and ct_s[-1].isdigit() and nt_s[0].isdigit():
                    self.logger.debug(f"[OOXMLRebuilder] Space rule override (digit->digit): seq_{current_seq_id} -> seq_{next_seq_id}, NO SPACE")
                    should_add = False
            if should_add:
                # 确保不重复添加空格
                if current_text and not current_text.endswith(' '):
                    original_text = current_text
                    current_segment['translated_text'] = current_text + ' '
                    spaces_added_count += 1
                    
                    # Debug: 记录空格添加详情
                    self.logger.debug(f"[OOXMLRebuilder] SPACE ADDED to segment seq_{current_seq_id}: {repr(original_text)} -> {repr(current_segment['translated_text'])}")
                else:
                    # Debug: 记录空格已存在的情况
                    if current_text.endswith(' '):
                        self.logger.debug(f"[OOXMLRebuilder] Space already exists in segment seq_{current_seq_id}, no addition needed")
                    else:
                        self.logger.debug(f"[OOXMLRebuilder] Empty current text in segment seq_{current_seq_id}, no space needed")
            else:
                # Debug: 记录不需要添加空格的情况
                self.logger.debug(f"[OOXMLRebuilder] No space needed between segment seq_{current_seq_id} and seq_{next_seq_id}")
        
        # Debug: 记录最终统计信息
        self.logger.debug(f"[OOXMLRebuilder] Space preprocessing completed:")
        self.logger.debug(f"[OOXMLRebuilder] - Total segments: {len(segments)}")
        self.logger.debug(f"[OOXMLRebuilder] - Processed pairs: {processed_pairs_count}")
        self.logger.debug(f"[OOXMLRebuilder] - Spaces added: {spaces_added_count}")
        self.logger.debug(f"[OOXMLRebuilder] - Space addition rate: {spaces_added_count/processed_pairs_count*100:.1f}%" if processed_pairs_count > 0 else "0.0%")
        
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
            # Debug: 记录开始XML替换
            self.logger.debug(f"[OOXMLRebuilder] Starting DOCX XML replacement for {len(segments)} segments")
            
            root = self._secure_parse_xml(xml_content)
            replaced_count = 0
            skipped_count = 0
            failed_count = 0
            
            # Sort segments by sequence_id to maintain order
            segments.sort(key=lambda x: x.get('sequence_id', 0))
            
            # Debug: 记录排序后的segments信息
            if segments:
                first_seq = segments[0].get('sequence_id', 'N/A')
                last_seq = segments[-1].get('sequence_id', 'N/A')
                self.logger.debug(f"[OOXMLRebuilder] DOCX segments sorted by sequence_id: {first_seq} -> {last_seq}")
            
            # Process segments with simplified replacement
            for idx, segment in enumerate(segments):
                seq_id = segment.get('sequence_id', 'N/A')
                
                # 获取完整的translated_text，包括空字符串（用于缺失翻译的替换）
                translated_text = segment.get('translated_text', '')
                if translated_text is None:  # 只跳过None值，保留空字符串用于替换原文
                    skipped_count += 1
                    self.logger.debug(f"[OOXMLRebuilder] DOCX segment {idx+1} (seq_{seq_id}): SKIPPED (None translated_text)")
                    continue
                
                # Use translated_text directly - spaces already added in preprocessing
                final_text = translated_text
                
                xpath = segment.get('xml_location', {}).get('element_xpath', '')
                namespace_map = segment.get('xml_location', {}).get('namespace_map', {})
                original_text = segment.get('original_text', '')
                
                # Debug: 记录替换前的信息
                self.logger.debug(f"[OOXMLRebuilder] DOCX segment {idx+1} (seq_{seq_id}): Processing replacement")
                self.logger.debug(f"[OOXMLRebuilder] - XPath: {xpath}")
                self.logger.debug(f"[OOXMLRebuilder] - Original: {repr(original_text)}")
                self.logger.debug(f"[OOXMLRebuilder] - Final: {repr(final_text)}")
                
                # 空格保持验证：检查final_text中的空格
                space_count = final_text.count(' ')
                ends_with_space = final_text.endswith(' ')
                self.logger.debug(f"[OOXMLRebuilder] - Space analysis: {space_count} spaces total, ends_with_space={ends_with_space}")
                
                if xpath:
                    try:
                        # Find the text element using xpath
                        elements = root.xpath(xpath, namespaces=namespace_map)
                        if elements:
                            # 记录替换前的元素内容
                            old_element_text = elements[0].text or ''
                            
                            # 执行替换
                            elements[0].text = final_text
                            # 保留首尾/连续空格
                            self._apply_xml_space_preserve(elements[0], final_text)
                            # 保留首尾/连续空格
                            self._apply_xml_space_preserve(elements[0], final_text)
                            # 保留首尾/连续空格
                            self._apply_xml_space_preserve(elements[0], final_text)
                            # 保留首尾/连续空格
                            self._apply_xml_space_preserve(elements[0], final_text)
                            # 保留首尾/连续空格
                            self._apply_xml_space_preserve(elements[0], final_text)
                            replaced_count += 1
                            
                            # Debug: 记录成功的替换
                            self.logger.debug(f"[OOXMLRebuilder] DOCX segment {idx+1} (seq_{seq_id}): REPLACED successfully")
                            self.logger.debug(f"[OOXMLRebuilder] - Element before: {repr(old_element_text)}")
                            self.logger.debug(f"[OOXMLRebuilder] - Element after: {repr(elements[0].text)}")
                            
                            # 空格保持验证：确认替换后空格是否保持
                            if ends_with_space and not elements[0].text.endswith(' '):
                                self.logger.warning(f"[OOXMLRebuilder] SPACE LOST during XML replacement for segment seq_{seq_id}!")
                            elif ends_with_space and elements[0].text.endswith(' '):
                                self.logger.debug(f"[OOXMLRebuilder] Space preserved in XML replacement for segment seq_{seq_id}")
                        else:
                            failed_count += 1
                            self.logger.warning(f"[OOXMLRebuilder] DOCX segment {idx+1} (seq_{seq_id}): FAILED - No elements found at xpath {xpath}")
                    except Exception as e:
                        failed_count += 1
                        self.logger.warning(f"[OOXMLRebuilder] DOCX segment {idx+1} (seq_{seq_id}): FAILED - Exception at xpath {xpath}: {str(e)}")
                else:
                    failed_count += 1
                    self.logger.warning(f"[OOXMLRebuilder] DOCX segment {idx+1} (seq_{seq_id}): FAILED - Empty xpath")
            
            # Debug: 记录最终统计信息
            total_segments = len(segments)
            self.logger.debug(f"[OOXMLRebuilder] DOCX XML replacement completed:")
            self.logger.debug(f"[OOXMLRebuilder] - Total segments: {total_segments}")
            self.logger.debug(f"[OOXMLRebuilder] - Replaced: {replaced_count}")
            self.logger.debug(f"[OOXMLRebuilder] - Skipped: {skipped_count}")
            self.logger.debug(f"[OOXMLRebuilder] - Failed: {failed_count}")
            self.logger.debug(f"[OOXMLRebuilder] - Success rate: {replaced_count/total_segments*100:.1f}%" if total_segments > 0 else "0.0%")
            
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
