import json
import logging
import xml.etree.ElementTree as ET
import re
import ast
from collections.abc import Generator
from typing import Any

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
            # Get parameters with dynamic type detection
            file_id = tool_parameters.get("file_id", "")
            translated_texts = tool_parameters.get("translated_texts", "")
            
            # Validate required parameters
            if not file_id:
                logger.error("[UpdateTranslations] Missing required parameter: file_id")
                yield self.create_text_message("Error: file_id is required.")
                return
            
            if not translated_texts:
                logger.error("[UpdateTranslations] Missing translation input: translated_texts is required")
                yield self.create_text_message("Error: translated_texts is required.")
                return
            
            # Dynamic type detection and processing with string-array parsing
            logger.info(f"[UpdateTranslations] Parameter type: {type(translated_texts)}")
            logger.debug(f"[UpdateTranslations] Parameter content preview: {str(translated_texts)[:200]}...")
            
            if isinstance(translated_texts, list):
                # Real array input format (from iteration node)
                input_format = "array_chunks"
                chunks_processed = len(translated_texts)
                combined_translations = self._combine_chunks(translated_texts)
                logger.info(f"[UpdateTranslations] Processing real array input - {chunks_processed} chunks")
                yield self.create_text_message(f"Processing {chunks_processed} translation chunks for file_id: {file_id}")
            
            elif isinstance(translated_texts, str) and translated_texts.startswith("['") and translated_texts.endswith("']"):
                # String representation of array from iteration node (serialized array)
                logger.info("[UpdateTranslations] Detected serialized array format - attempting to parse")
                try:
                    parsed_array = ast.literal_eval(translated_texts)
                    if isinstance(parsed_array, list):
                        input_format = "array_chunks"  
                        chunks_processed = len(parsed_array)
                        combined_translations = self._combine_chunks(parsed_array)
                        logger.info(f"[UpdateTranslations] Successfully parsed serialized array - {chunks_processed} chunks")
                        yield self.create_text_message(f"Processing {chunks_processed} parsed translation chunks for file_id: {file_id}")
                    else:
                        # Parsed result is not a list, fallback to string processing
                        input_format = "single_string"
                        chunks_processed = 0
                        combined_translations = str(translated_texts)
                        logger.warning("[UpdateTranslations] Parsed result is not a list - falling back to string processing")
                        yield self.create_text_message(f"Processing single translation string for file_id: {file_id}")
                except (ValueError, SyntaxError) as e:
                    # Parse failed, treat as regular string
                    input_format = "single_string"
                    chunks_processed = 0
                    combined_translations = str(translated_texts)
                    logger.warning(f"[UpdateTranslations] Failed to parse serialized array ({str(e)}) - treating as string")
                    yield self.create_text_message(f"Processing single translation string for file_id: {file_id}")
            
            else:
                # Single string input format (original workflow)
                input_format = "single_string"
                chunks_processed = 0
                combined_translations = str(translated_texts)  # Ensure string type
                logger.info(f"[UpdateTranslations] Processing single string input - {len(combined_translations)} characters")
                yield self.create_text_message(f"Processing single translation string for file_id: {file_id}")
            
            logger.info(f"[UpdateTranslations] Input format detected: {input_format}")
            logger.debug(f"[UpdateTranslations] Combined translation length: {len(combined_translations)} characters")
            
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
            
            # 解析XML格式的翻译结果，处理重叠冲突
            logger.debug(f"[UpdateTranslations] Parsing XML formatted translations")
            translation_dict = self._parse_xml_translations(combined_translations)
            logger.info(f"[UpdateTranslations] Found {len(translation_dict)} translation segments")
            
            # Debug: Show sample translations
            if translation_dict:
                sample_items = list(translation_dict.items())[:3]
                logger.debug(f"[UpdateTranslations] Sample translations: {[(k, v[:50] + '...' if len(v) > 50 else v) for k, v in sample_items]}")
            
            # 读取映射关系
            logger.debug(f"[UpdateTranslations] Reading segment mapping")
            segment_mapping = self._get_segment_mapping(file_id)
            logger.debug(f"[UpdateTranslations] Mapping loaded - {len(segment_mapping)} mappings found")
            
            # 统计预期的翻译数量
            expected_translations = len(segment_mapping) if segment_mapping else len(text_segments)
            logger.info(f"[UpdateTranslations] Expected translations: {expected_translations}")
            
            # 检查翻译数量匹配
            actual_translations = len(translation_dict)
            mismatch_warning = expected_translations != actual_translations
            
            if mismatch_warning:
                logger.warning(f"[UpdateTranslations] Count mismatch: expected {expected_translations}, got {actual_translations}")
                yield self.create_text_message(
                    f"Warning: Translation count mismatch. "
                    f"Expected {expected_translations} translations, got {actual_translations}"
                )
            
            # 简化翻译更新：直接使用翻译结果，不做任何空格处理
            logger.info(f"[UpdateTranslations] Starting simplified translation updates")
            updated_count = 0
            skipped_count = 0
            
            for i, segment in enumerate(text_segments):
                if str(i) in segment_mapping:
                    xml_id = segment_mapping[str(i)]
                    if xml_id in translation_dict:
                        translation_text = translation_dict[xml_id]
                        if translation_text:
                            # 关键简化：直接使用翻译结果，不做任何空格处理
                            segment['translated_text'] = translation_text
                            updated_count += 1
                            logger.debug(f"[UpdateTranslations] Updated index {i}, xml_id {xml_id}: {translation_text[:50]}...") if updated_count <= 5 else None
                        else:
                            skipped_count += 1
                            logger.debug(f"[UpdateTranslations] Skipped empty translation for index {i}, xml_id {xml_id}")
                    else:
                        skipped_count += 1
                        logger.warning(f"[UpdateTranslations] Missing translation for index {i}, xml_id {xml_id}")
                elif segment_mapping:
                    # 有映射但当前segment不在映射中，跳过
                    skipped_count += 1
                    logger.debug(f"[UpdateTranslations] No mapping found for segment at index {i}")
                else:
                    # 兼容旧格式：没有映射关系时使用简单的顺序映射
                    xml_id = f"{i + 1:03d}"
                    if xml_id in translation_dict:
                        translation_text = translation_dict[xml_id]
                        if translation_text:
                            segment['translated_text'] = translation_text
                            updated_count += 1
                            logger.debug(f"[UpdateTranslations] Legacy update - index {i}, xml_id {xml_id}: {translation_text[:50]}...") if updated_count <= 5 else None
                        else:
                            skipped_count += 1
                    else:
                        skipped_count += 1
                        logger.warning(f"[UpdateTranslations] Legacy mode - missing translation for index {i}, xml_id {xml_id}")
            
            # Store updated text segments
            logger.info(f"[UpdateTranslations] Storing updated text segments")
            storage_start_time = self._get_current_timestamp()
            self._store_text_segments(file_id, text_segments)
            storage_end_time = self._get_current_timestamp()
            logger.debug(f"[UpdateTranslations] Storage completed - Start: {storage_start_time}, End: {storage_end_time}")
            
            logger.info(f"[UpdateTranslations] Update completed successfully - Updated: {updated_count}, Skipped: {skipped_count}, Mismatch: {mismatch_warning}")
            yield self.create_text_message(
                f"Updated {updated_count} translations, skipped {skipped_count} entries"
            )
            
            # Return detailed results
            result = {
                "success": True,
                "file_id": file_id,
                "updated_count": updated_count,
                "skipped_count": skipped_count,
                "mismatch_warning": mismatch_warning,
                "chunks_processed": chunks_processed,
                "input_format": input_format,
                "message": f"Successfully updated {updated_count} translations using {input_format} input"
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
            
            # 处理重叠ID冲突：后面的翻译优先
            translation_dict = self._resolve_overlap_conflicts(translation_dict)
            logger.debug(f"[UpdateTranslations] Overlap conflicts resolved - Final count: {len(translation_dict)}")
            
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
        """Get text segments from persistent storage with simplified logic."""
        logger.debug(f"[UpdateTranslations] Reading text segments for file_id: {file_id}")
        try:
            texts_key = f"{file_id}_texts"
            logger.debug(f"[UpdateTranslations] Fetching text segments with key: {texts_key}")
            texts_data = self.session.storage.get(texts_key)
            if not texts_data:
                logger.warning(f"[UpdateTranslations] No text segments found for key: {texts_key}")
                return []
            
            # Parse JSON segments directly
            if isinstance(texts_data, bytes):
                segments = json.loads(texts_data.decode('utf-8'))
            else:
                segments = json.loads(texts_data)
            logger.debug(f"[UpdateTranslations] Text segments loaded successfully - Count: {len(segments)}")
            
            # Log translation status for debugging
            translated_count = sum(1 for seg in segments if seg.get('translated_text', '').strip())
            logger.debug(f"[UpdateTranslations] Existing translations found: {translated_count}/{len(segments)}")
            
            return segments
        except Exception as e:
            logger.error(f"[UpdateTranslations] Error reading text segments: {str(e)}")
            return []
    
    def _store_text_segments(self, file_id: str, text_segments: list):
        """Store updated text segments to persistent storage with simplified logic."""
        logger.debug(f"[UpdateTranslations] Storing text segments for file_id: {file_id}, count: {len(text_segments)}")
        try:
            # Store segments directly with proper bytes encoding
            texts_key = f"{file_id}_texts"
            self.session.storage.set(texts_key, json.dumps(text_segments, ensure_ascii=False).encode('utf-8'))
            
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
    
    
    
    
    def _get_metadata(self, file_id: str) -> dict:
        """Get metadata from persistent storage with simplified logic."""
        logger.debug(f"[UpdateTranslations] Reading metadata for file_id: {file_id}")
        try:
            metadata_key = f"{file_id}_metadata"
            logger.debug(f"[UpdateTranslations] Fetching metadata with key: {metadata_key}")
            metadata_data = self.session.storage.get(metadata_key)
            if not metadata_data:
                logger.warning(f"[UpdateTranslations] No metadata found for key: {metadata_key}")
                return {}
            
            # Parse JSON metadata directly
            if isinstance(metadata_data, bytes):
                metadata = json.loads(metadata_data.decode('utf-8'))
            else:
                metadata = json.loads(metadata_data)
            logger.debug(f"[UpdateTranslations] Metadata loaded successfully - Keys: {list(metadata.keys())}")
            
            return metadata
        except Exception as e:
            logger.error(f"[UpdateTranslations] Error reading metadata: {str(e)}")
            return {}
    
    def _combine_chunks(self, chunks: list) -> str:
        """
        将分块的XML翻译结果合并为单个字符串进行处理。
        
        Args:
            chunks: XML翻译分块列表
            
        Returns:
            合并后的XML字符串
        """
        if not chunks:
            return ""
        
        logger.debug(f"[UpdateTranslations] Combining {len(chunks)} translation chunks")
        
        try:
            # 验证chunks是否为字符串列表
            if not all(isinstance(chunk, str) for chunk in chunks):
                logger.warning("[UpdateTranslations] Some chunks are not strings, converting to string")
                chunks = [str(chunk) for chunk in chunks]
            
            # 合并所有chunks，使用换行符分隔
            combined = '\n'.join(chunk.strip() for chunk in chunks if chunk and chunk.strip())
            
            # 计算统计信息
            total_length = len(combined)
            non_empty_chunks = len([chunk for chunk in chunks if chunk and chunk.strip()])
            
            logger.info(f"[UpdateTranslations] Chunks combined successfully - {non_empty_chunks}/{len(chunks)} non-empty chunks, {total_length} total characters")
            logger.debug(f"[UpdateTranslations] Combined content preview: {combined[:200]}...")
            
            return combined
            
        except Exception as e:
            logger.error(f"[UpdateTranslations] Error combining chunks: {str(e)}")
            # 容错处理：返回空字符串
            return ""
    
    def _resolve_overlap_conflicts(self, translation_dict: dict) -> dict:
        """
        解决重叠分块导致的ID冲突，按照'后一部分优先'原则。
        
        Args:
            translation_dict: 原始翻译字典，可能包含重复ID
            
        Returns:
            解决冲突后的翻译字典
        """
        if not translation_dict:
            return translation_dict
        
        # 检查是否有重复的ID（重叠分块导致）
        id_counts = {}
        for segment_id in translation_dict.keys():
            id_counts[segment_id] = id_counts.get(segment_id, 0) + 1
        
        duplicates = [id for id, count in id_counts.items() if count > 1]
        
        if not duplicates:
            logger.debug(f"[UpdateTranslations] No overlap conflicts found")
            return translation_dict
        
        logger.info(f"[UpdateTranslations] Found {len(duplicates)} overlapping IDs: {duplicates[:5]}...")
        
        # 由于字典的特性，后面的值会自动覆盖前面的值
        # 这正好符合"后一部分优先"的原则
        resolved_dict = {}
        conflict_count = 0
        
        for segment_id, translation in translation_dict.items():
            if segment_id in duplicates:
                if segment_id in resolved_dict:
                    conflict_count += 1
                    logger.debug(f"[UpdateTranslations] Overlap resolved for ID {segment_id}: using later translation")
                resolved_dict[segment_id] = translation  # 后面的翻译覆盖前面的
            else:
                resolved_dict[segment_id] = translation
        
        logger.info(f"[UpdateTranslations] Overlap conflicts resolved - {conflict_count} conflicts resolved using later translations")
        return resolved_dict