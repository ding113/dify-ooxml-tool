import json
import gzip
import logging
import xml.etree.ElementTree as ET
import re
from collections.abc import Generator
from typing import Any

# Use orjson for high-performance JSON deserialization, fallback to standard json
try:
    import orjson
    ORJSON_AVAILABLE = True
except ImportError:
    ORJSON_AVAILABLE = False

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.config.logger_format import plugin_logger_handler

# Initialize logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
logger.addHandler(plugin_logger_handler)

class UpdateTranslationsTool(Tool):
    """
    Update the stored texts with LLM translation results.
    
    This tool receives translated texts from LLM and updates the stored text segments
    while maintaining exact order correspondence with the original texts.
    """
    
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage, None, None]:
        logger.info("[UpdateTranslations] Starting translation update process")
        try:
            # Get parameters
            file_id = tool_parameters.get("file_id", "")
            translated_texts = tool_parameters.get("translated_texts", "")
            
            # Validate required parameters
            if not file_id:
                logger.error("[UpdateTranslations] Missing required parameter: file_id")
                yield self.create_text_message("Error: file_id is required.")
                return
            
            if not translated_texts:
                logger.error("[UpdateTranslations] Missing required parameter: translated_texts")
                yield self.create_text_message("Error: translated_texts is required.")
                return
            
            logger.info(f"[UpdateTranslations] Processing file_id: {file_id}")
            logger.debug(f"[UpdateTranslations] Translation input length: {len(translated_texts)} characters")
            yield self.create_text_message(f"Updating translations for file_id: {file_id}")
            
            # Read existing text segments
            logger.info(f"[UpdateTranslations] Reading existing text segments")
            text_segments = self._get_text_segments(file_id)
            if not text_segments:
                logger.error("[UpdateTranslations] No text segments found")
                yield self.create_text_message("Error: No text segments found for this file_id. Please run extract_ooxml_text first.")
                return
            logger.info(f"[UpdateTranslations] Loaded {len(text_segments)} text segments from storage")
            
            # Sort by sequence_id to ensure correct order
            logger.debug(f"[UpdateTranslations] Sorting {len(text_segments)} segments by sequence_id")
            text_segments.sort(key=lambda x: x.get('sequence_id', 0))
            logger.debug(f"[UpdateTranslations] Sorting completed - First segment ID: {text_segments[0].get('sequence_id', 'N/A') if text_segments else 'N/A'}")
            
            # Parse the XML formatted translated texts
            logger.debug(f"[UpdateTranslations] Parsing XML formatted translations")
            translation_dict = self._parse_xml_translations(translated_texts)
            logger.info(f"[UpdateTranslations] Found {len(translation_dict)} translation segments")
            
            # Debug: Show sample translations
            if translation_dict:
                sample_items = list(translation_dict.items())[:3]
                logger.debug(f"[UpdateTranslations] Sample translations: {[(k, v[:50] + '...' if len(v) > 50 else v) for k, v in sample_items]}")
            
            # 读取映射关系（用于新格式的空格恢复）
            logger.debug(f"[UpdateTranslations] Reading segment mapping for space recovery")
            segment_mapping = self._get_segment_mapping(file_id)
            reverse_mapping = {v: k for k, v in segment_mapping.items()} if segment_mapping else {}
            logger.debug(f"[UpdateTranslations] Mapping loaded - {len(segment_mapping)} mappings found")
            
            # 统计预期的翻译数量
            if segment_mapping:
                # 新格式：使用映射关系
                expected_translations = len(segment_mapping)
                logger.info(f"[UpdateTranslations] Using new mapping format - Expected translations: {expected_translations}")
            else:
                # 旧格式：兼容处理
                non_empty_indices = []
                for i, segment in enumerate(text_segments):
                    if segment.get('original_text', '').strip():
                        non_empty_indices.append(i)
                expected_translations = len(non_empty_indices)
                logger.info(f"[UpdateTranslations] Using legacy format - Expected translations: {expected_translations}")
            
            # 检查翻译数量匹配
            actual_translations = len(translation_dict)
            mismatch_warning = expected_translations != actual_translations
            
            if mismatch_warning:
                logger.warning(f"[UpdateTranslations] Count mismatch: expected {expected_translations}, got {actual_translations}")
                yield self.create_text_message(
                    f"Warning: Translation count mismatch. "
                    f"Expected {expected_translations} translations, got {actual_translations}"
                )
            
            # 智能更新翻译，实现空格恢复
            logger.info(f"[UpdateTranslations] Starting intelligent translation updates with space recovery")
            updated_count = 0
            skipped_count = 0
            space_recovered_count = 0
            
            for i, segment in enumerate(text_segments):
                # 获取空格信息（新格式）或兼容处理（旧格式）
                space_info = segment.get('space_info', {})
                
                if space_info:
                    # 新格式：使用space_info
                    has_content = space_info.get('has_content', True)
                    is_pure_space = space_info.get('is_pure_space', False)
                    raw_text = space_info.get('raw_text', '')
                    leading_spaces = space_info.get('leading_spaces', 0)
                    trailing_spaces = space_info.get('trailing_spaces', 0)
                    
                    if is_pure_space:
                        # 纯空格段落：直接使用原始文本
                        segment['translated_text'] = raw_text
                        space_recovered_count += 1
                        logger.debug(f"[UpdateTranslations] Space recovered at index {i}: {repr(raw_text)}") if space_recovered_count <= 3 else None
                        continue
                    elif has_content and str(i) in segment_mapping:
                        # 有内容的段落：查找翻译并恢复空格
                        xml_id = segment_mapping[str(i)]
                        if xml_id in translation_dict:
                            translation_text = translation_dict[xml_id].strip()
                            if translation_text:
                                # 恢复前后空格
                                leading = ' ' * leading_spaces
                                trailing = ' ' * trailing_spaces
                                full_translation = leading + translation_text + trailing
                                segment['translated_text'] = full_translation
                                updated_count += 1
                                logger.debug(f"[UpdateTranslations] Updated with spaces - index {i}, xml_id {xml_id}: {repr(full_translation[:50])}") if updated_count <= 5 else None
                            else:
                                skipped_count += 1
                                logger.debug(f"[UpdateTranslations] Skipped empty translation for index {i}, xml_id {xml_id}")
                        else:
                            skipped_count += 1
                            logger.warning(f"[UpdateTranslations] Missing translation for index {i}, xml_id {xml_id}")
                    elif has_content:
                        # 有内容但未找到映射：使用原始文本
                        segment['translated_text'] = raw_text
                        skipped_count += 1
                        logger.debug(f"[UpdateTranslations] No mapping found for content segment at index {i}, using original")
                else:
                    # 旧格式：兼容处理
                    original_text = segment.get('original_text', '').strip()
                    if original_text:
                        # 寻找对应的翻译（使用简单的顺序映射）
                        xml_id_index = len([s for j, s in enumerate(text_segments[:i]) if s.get('original_text', '').strip()])
                        xml_id = f"{xml_id_index + 1:03d}"
                        
                        if xml_id in translation_dict:
                            translation_text = translation_dict[xml_id].strip()
                            if translation_text:
                                segment['translated_text'] = translation_text
                                updated_count += 1
                                logger.debug(f"[UpdateTranslations] Legacy update - index {i}, xml_id {xml_id}: {translation_text[:50]}...") if updated_count <= 5 else None
                            else:
                                skipped_count += 1
                        else:
                            skipped_count += 1
                            logger.warning(f"[UpdateTranslations] Legacy mode - missing translation for index {i}, xml_id {xml_id}")
                    # 空白段落在旧格式中保持不变
            
            # Store updated text segments
            logger.info(f"[UpdateTranslations] Storing updated text segments")
            storage_start_time = self._get_current_timestamp()
            self._store_text_segments(file_id, text_segments)
            storage_end_time = self._get_current_timestamp()
            logger.debug(f"[UpdateTranslations] Storage completed - Start: {storage_start_time}, End: {storage_end_time}")
            
            logger.info(f"[UpdateTranslations] Update completed successfully - Updated: {updated_count}, Spaces recovered: {space_recovered_count}, Skipped: {skipped_count}, Mismatch: {mismatch_warning}")
            yield self.create_text_message(
                f"Updated {updated_count} translations, recovered {space_recovered_count} space segments, skipped {skipped_count} entries"
            )
            
            # Return detailed results
            result = {
                "success": True,
                "file_id": file_id,
                "updated_count": updated_count,
                "space_recovered_count": space_recovered_count,
                "skipped_count": skipped_count,
                "mismatch_warning": mismatch_warning,
                "message": f"Successfully updated {updated_count} translations and recovered {space_recovered_count} space segments"
            }
            
            yield self.create_json_message(result)
            yield self.create_variable_message("update_result", json.dumps(result))
            
        except Exception as e:
            error_msg = f"Failed to update translations: {str(e)}"
            logger.error(f"[UpdateTranslations] Exception occurred: {error_msg}")
            yield self.create_text_message(f"Error: {error_msg}")
            yield self.create_json_message({
                "success": False,
                "file_id": tool_parameters.get("file_id", ""),
                "error": error_msg
            })
    
    def _parse_xml_translations(self, xml_content: str) -> dict:
        """解析XML格式的翻译结果，返回ID到翻译文本的映射字典。"""
        logger.debug(f"[UpdateTranslations] Starting XML translation parsing")
        translation_dict = {}
        
        try:
            # 1. 首先移除think标签
            xml_content = self._remove_think_tags(xml_content)
            logger.debug(f"[UpdateTranslations] Think tags filtering completed")
            
            # 2. 清理XML内容
            xml_content = xml_content.strip()
            
            # 添加根元素包装（如果没有的话）
            if not xml_content.startswith('<root>'):
                xml_content = f'<root>{xml_content}</root>'
                logger.debug(f"[UpdateTranslations] Added root wrapper to XML content")
            
            # 解析XML
            root = ET.fromstring(xml_content)
            logger.debug(f"[UpdateTranslations] XML parsed successfully")
            
            # 查找所有segment元素
            segments = root.findall('.//segment')
            logger.debug(f"[UpdateTranslations] Found {len(segments)} segment elements")
            
            for segment in segments:
                segment_id = segment.get('id')
                text_content = segment.text or ''
                
                # XML反转义
                text_content = self._xml_unescape(text_content)
                
                if segment_id and text_content.strip():
                    translation_dict[segment_id] = text_content.strip()
                    logger.debug(f"[UpdateTranslations] Extracted segment {segment_id}: {text_content[:50]}...") if len(translation_dict) <= 5 else None
                else:
                    logger.warning(f"[UpdateTranslations] Invalid segment: id={segment_id}, text={text_content[:50] if text_content else 'empty'}")
            
            logger.info(f"[UpdateTranslations] XML parsing completed successfully - {len(translation_dict)} valid translations")
            return translation_dict
            
        except ET.ParseError as e:
            logger.warning(f"[UpdateTranslations] XML parsing failed: {str(e)}")
            logger.debug(f"[UpdateTranslations] XML content preview: {xml_content[:200]}...")
            # 容错处理：尝试正则表达式提取
            logger.info(f"[UpdateTranslations] Attempting regex fallback parsing")
            translation_dict = self._regex_fallback_parse(xml_content)
            logger.info(f"[UpdateTranslations] Regex fallback completed - {len(translation_dict)} translations extracted")
            return translation_dict
        
        except Exception as e:
            logger.error(f"[UpdateTranslations] Unexpected error in XML parsing: {str(e)}")
            return {}
    
    def _xml_unescape(self, text: str) -> str:
        """XML反转义，将XML实体转换回原始字符。"""
        if not text:
            return text
        
        try:
            # XML实体反转义
            unescaped = text.replace('&lt;', '<')
            unescaped = unescaped.replace('&gt;', '>')
            unescaped = unescaped.replace('&quot;', '"')
            unescaped = unescaped.replace('&apos;', "'")
            unescaped = unescaped.replace('&amp;', '&')  # 必须最后处理&
            
            return unescaped
        except Exception as e:
            logger.error(f"[UpdateTranslations] Error XML unescaping text: {str(e)}")
            return text  # Return original if unescaping fails
    
    def _remove_think_tags(self, content: str) -> str:
        """移除LLM输出中的<think></think>标签及其内容。
        
        Args:
            content: 原始LLM输出内容
            
        Returns:
            清理后的内容，已移除所有think标签
        """
        if not content:
            return content
        
        try:
            original_length = len(content)
            
            # 移除所有<think></think>标签及其内容
            # re.DOTALL: 让.匹配包括换行符的任意字符
            # re.IGNORECASE: 大小写不敏感
            # .*?: 非贪婪匹配，避免匹配过多内容
            pattern = r'<think>.*?</think>'
            cleaned_content = re.sub(pattern, '', content, flags=re.DOTALL | re.IGNORECASE)
            
            # 清理多余的空行和空白
            cleaned_content = re.sub(r'\n\s*\n', '\n', cleaned_content).strip()
            
            cleaned_length = len(cleaned_content)
            think_tags_found = original_length != cleaned_length
            
            if think_tags_found:
                logger.info(f"[UpdateTranslations] Think tags removed - Original: {original_length} chars, Cleaned: {cleaned_length} chars")
                logger.debug(f"[UpdateTranslations] Content preview after cleaning: {cleaned_content[:200]}...")
            else:
                logger.debug(f"[UpdateTranslations] No think tags found in content")
            
            return cleaned_content
            
        except Exception as e:
            logger.warning(f"[UpdateTranslations] Error removing think tags: {str(e)}")
            return content  # 返回原内容作为容错
    
    def _regex_fallback_parse(self, content: str) -> dict:
        """正则表达式容错解析，当XML解析失败时使用。"""
        logger.debug(f"[UpdateTranslations] Starting regex fallback parsing")
        translation_dict = {}
        
        try:
            # 正则模式匹配 <segment id="XXX">content</segment>
            pattern = r'<segment\s+id=["\']([^"\'>]+)["\']>([^<]*)</segment>'
            
            matches = re.findall(pattern, content, re.DOTALL | re.IGNORECASE)
            logger.debug(f"[UpdateTranslations] Regex found {len(matches)} segment matches")
            
            for segment_id, text_content in matches:
                if text_content.strip():
                    # XML反转义
                    clean_text = self._xml_unescape(text_content.strip())
                    translation_dict[segment_id] = clean_text
                    logger.debug(f"[UpdateTranslations] Regex extracted segment {segment_id}: {clean_text[:50]}...") if len(translation_dict) <= 5 else None
            
            logger.info(f"[UpdateTranslations] Regex fallback parsing completed - {len(translation_dict)} valid translations")
            return translation_dict
            
        except Exception as e:
            logger.error(f"[UpdateTranslations] Error in regex fallback parsing: {str(e)}")
            return {}
    
    def _get_text_segments(self, file_id: str) -> list:
        """Get text segments from persistent storage with compression and batch support."""
        logger.debug(f"[UpdateTranslations] Reading text segments for file_id: {file_id}")
        try:
            # First, check if we have batched storage
            batch_metadata = self._get_batch_metadata(file_id)
            if batch_metadata:
                return self._get_segments_batched(file_id, batch_metadata)
            
            # Try single storage (both compressed and legacy)
            texts_key = f"{file_id}_texts"
            logger.debug(f"[UpdateTranslations] Fetching text segments with key: {texts_key}")
            texts_data = self.session.storage.get(texts_key)
            if not texts_data:
                logger.warning(f"[UpdateTranslations] No text segments found for key: {texts_key}")
                return []
            
            # Try compressed format first (new format)
            try:
                decompressed_data = gzip.decompress(texts_data)
                segments = self._fast_json_decode(decompressed_data)
                logger.debug(f"[UpdateTranslations] Compressed text segments loaded successfully - Count: {len(segments)}")
            except (gzip.BadGzipFile, OSError):
                # Fallback to legacy format (uncompressed)
                segments = json.loads(texts_data.decode('utf-8'))
                logger.debug(f"[UpdateTranslations] Legacy text segments loaded successfully - Count: {len(segments)}")
            
            # Restore namespace maps if optimized format is detected
            segments = self._restore_optimized_data(segments, file_id)
            
            # Log translation status for debugging
            translated_count = sum(1 for seg in segments if seg.get('translated_text', '').strip())
            logger.debug(f"[UpdateTranslations] Existing translations found: {translated_count}/{len(segments)}")
            
            return segments
        except Exception as e:
            logger.error(f"[UpdateTranslations] Error reading text segments: {str(e)}")
            return []
    
    def _store_text_segments(self, file_id: str, text_segments: list):
        """Store updated text segments to persistent storage."""
        logger.debug(f"[UpdateTranslations] Storing text segments for file_id: {file_id}, count: {len(text_segments)}")
        try:
            texts_key = f"{file_id}_texts"
            # Optimize JSON serialization - do it once
            texts_bytes = json.dumps(text_segments, ensure_ascii=False).encode('utf-8')
            json_size = len(texts_bytes)
            logger.debug(f"[UpdateTranslations] Storing segments - Key: {texts_key}, Size: {json_size} bytes")
            self.session.storage.set(texts_key, texts_bytes)
            
            # Verify storage by counting translations
            translated_count = sum(1 for seg in text_segments if seg.get('translated_text', '').strip())
            logger.debug(f"[UpdateTranslations] Storage verification - Total segments: {len(text_segments)}, Translated: {translated_count}")
            logger.info(f"[UpdateTranslations] Text segments stored successfully")
        except Exception as e:
            logger.error(f"[UpdateTranslations] Failed to store text segments: {str(e)}")
            logger.debug(f"[UpdateTranslations] Storage error details - segments_count: {len(text_segments)}, file_id: {file_id}")
            raise Exception(f"Failed to store updated text segments: {str(e)}")    
    
    def _get_current_timestamp(self) -> str:
        """Get current timestamp in ISO format."""
        from datetime import datetime
        return datetime.now().isoformat()
    
    def _get_segment_mapping(self, file_id: str) -> dict:
        """读取segment映射关系，用于空格恢复，支持压缩格式。"""
        logger.debug(f"[UpdateTranslations] Reading segment mapping for file_id: {file_id}")
        try:
            mapping_key = f"{file_id}_mapping"
            logger.debug(f"[UpdateTranslations] Fetching mapping with key: {mapping_key}")
            mapping_data = self.session.storage.get(mapping_key)
            if not mapping_data:
                logger.warning(f"[UpdateTranslations] No mapping found for key: {mapping_key} (legacy format)")
                return {}
            
            # Try compressed format first (new format)
            try:
                decompressed_data = gzip.decompress(mapping_data)
                mapping = self._fast_json_decode(decompressed_data)
                logger.debug(f"[UpdateTranslations] Compressed mapping loaded successfully - Count: {len(mapping)}")
            except (gzip.BadGzipFile, OSError):
                # Fallback to legacy format (uncompressed)
                mapping = json.loads(mapping_data.decode('utf-8'))
                logger.debug(f"[UpdateTranslations] Legacy mapping loaded successfully - Count: {len(mapping)}")
            
            return mapping
        except Exception as e:
            logger.error(f"[UpdateTranslations] Error reading segment mapping: {str(e)}")
            return {}
    
    def _fast_json_decode(self, data: bytes) -> any:
        """High-performance JSON decoding with orjson fallback."""
        if ORJSON_AVAILABLE:
            return orjson.loads(data)
        else:
            return json.loads(data.decode('utf-8'))
    
    def _get_batch_metadata(self, file_id: str) -> dict:
        """Get batch metadata for batched storage format."""
        try:
            batch_meta_key = f"{file_id}_batch_metadata"
            batch_meta_data = self.session.storage.get(batch_meta_key)
            if not batch_meta_data:
                return {}
            
            # Try compressed format first
            try:
                decompressed_data = gzip.decompress(batch_meta_data)
                return self._fast_json_decode(decompressed_data)
            except (gzip.BadGzipFile, OSError):
                return self._fast_json_decode(batch_meta_data)
        except Exception as e:
            logger.debug(f"[UpdateTranslations] No batch metadata found: {str(e)}")
            return {}
    
    def _get_segments_batched(self, file_id: str, batch_metadata: dict) -> list:
        """Retrieve text segments from batched storage format."""
        batch_count = batch_metadata.get('batch_count', 0)
        total_segments = batch_metadata.get('total_segments', 0)
        
        logger.debug(f"[UpdateTranslations] Loading batched segments: {batch_count} batches, {total_segments} total segments")
        
        all_segments = []
        for batch_index in range(batch_count):
            batch_key = f"{file_id}_texts_batch_{batch_index}"
            batch_data = self.session.storage.get(batch_key)
            
            if batch_data:
                try:
                    # Try compressed format first
                    try:
                        decompressed_data = gzip.decompress(batch_data)
                        batch_segments = self._fast_json_decode(decompressed_data)
                    except (gzip.BadGzipFile, OSError):
                        batch_segments = self._fast_json_decode(batch_data)
                    
                    all_segments.extend(batch_segments)
                    logger.debug(f"[UpdateTranslations] Loaded batch {batch_index}: {len(batch_segments)} segments")
                except Exception as e:
                    logger.warning(f"[UpdateTranslations] Failed to load batch {batch_index}: {str(e)}")
            else:
                logger.warning(f"[UpdateTranslations] Batch {batch_index} not found")
        
        logger.info(f"[UpdateTranslations] Batched loading complete: {len(all_segments)}/{total_segments} segments loaded")
        
        # Restore namespace maps if optimized format is detected
        return self._restore_optimized_data(all_segments, file_id)
    
    def _restore_optimized_data(self, segments: list, file_id: str) -> list:
        """Restore namespace maps from optimized format using metadata registry."""
        if not segments:
            return segments
        
        # Check if any segment has namespace_ref (indicating optimized format)
        has_optimization = any(
            seg.get('xml_location', {}).get('namespace_ref') is not None
            for seg in segments
        )
        
        if not has_optimization:
            logger.debug("[UpdateTranslations] No optimization detected, segments unchanged")
            return segments
        
        # Get namespace registry from metadata
        metadata = self._get_metadata(file_id)
        namespace_registry = metadata.get('namespace_registry', {})
        
        if not namespace_registry:
            logger.warning("[UpdateTranslations] Optimized format detected but no namespace registry found")
            return segments
        
        # Restore namespace maps
        restored_segments = []
        for segment in segments:
            restored_segment = segment.copy()
            xml_location = restored_segment.get('xml_location', {})
            namespace_ref = xml_location.get('namespace_ref')
            
            if namespace_ref and namespace_ref in namespace_registry:
                # Restore the full namespace_map
                xml_location['namespace_map'] = namespace_registry[namespace_ref]
                xml_location.pop('namespace_ref', None)  # Remove the reference
            
            restored_segments.append(restored_segment)
        
        logger.debug(f"[UpdateTranslations] Data restoration complete: {len(namespace_registry)} namespace patterns restored")
        return restored_segments
    
    def _get_metadata(self, file_id: str) -> dict:
        """Get metadata from persistent storage with compression support."""
        logger.debug(f"[UpdateTranslations] Reading metadata for file_id: {file_id}")
        try:
            metadata_key = f"{file_id}_metadata"
            logger.debug(f"[UpdateTranslations] Fetching metadata with key: {metadata_key}")
            metadata_data = self.session.storage.get(metadata_key)
            if not metadata_data:
                logger.warning(f"[UpdateTranslations] No metadata found for key: {metadata_key}")
                return {}
            
            # Try compressed format first (new format)
            try:
                decompressed_data = gzip.decompress(metadata_data)
                metadata = self._fast_json_decode(decompressed_data)
                logger.debug(f"[UpdateTranslations] Compressed metadata loaded successfully - Keys: {list(metadata.keys())}")
            except (gzip.BadGzipFile, OSError):
                # Fallback to legacy format (uncompressed)
                metadata = json.loads(metadata_data.decode('utf-8'))
                logger.debug(f"[UpdateTranslations] Legacy metadata loaded successfully - Keys: {list(metadata.keys())}")
            
            return metadata
        except Exception as e:
            logger.error(f"[UpdateTranslations] Error reading metadata: {str(e)}")
            return {}