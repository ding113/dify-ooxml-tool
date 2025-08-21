import json
import logging
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.config.logger_format import plugin_logger_handler

# Initialize logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
logger.addHandler(plugin_logger_handler)

class GetTranslationTextsTool(Tool):
    """
    Retrieve all extracted texts in order for LLM translation.
    
    This tool reads the previously extracted text segments from persistent storage
    and formats them for LLM translation while preserving the exact order.
    """
    
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage, None, None]:
        logger.info("[GetTranslationTexts] Starting text retrieval for translation")
        try:
            # Get parameters
            file_id = tool_parameters.get("file_id", "")
            
            # Validate required parameters
            if not file_id:
                logger.error("[GetTranslationTexts] Missing required parameter: file_id")
                yield self.create_text_message("Error: file_id is required.")
                return
            
            logger.info(f"[GetTranslationTexts] Processing file_id: {file_id}")
            logger.debug(f"[GetTranslationTexts] Starting XML text retrieval process")
            yield self.create_text_message(f"Retrieving texts for translation from file_id: {file_id}")
            
            # Read metadata
            logger.info(f"[GetTranslationTexts] Reading metadata")
            metadata = self._get_metadata(file_id)
            if not metadata:
                logger.error("[GetTranslationTexts] No metadata found")
                yield self.create_text_message("Error: No extraction data found for this file_id. Please run extract_ooxml_text first.")
                return
            file_type = metadata.get('file_type', 'unknown')
            total_count = metadata.get('total_text_count', 0)
            logger.info(f"[GetTranslationTexts] Metadata loaded - File type: {file_type}, Expected text count: {total_count}")
            
            # Read text segments
            logger.info(f"[GetTranslationTexts] Reading text segments")
            text_segments = self._get_text_segments(file_id)
            if not text_segments:
                logger.error("[GetTranslationTexts] No text segments found")
                yield self.create_text_message("Error: No text segments found for this file_id.")
                return
            logger.info(f"[GetTranslationTexts] Loaded {len(text_segments)} text segments from storage")
            
            # Sort by sequence_id to ensure correct order
            logger.debug(f"[GetTranslationTexts] Sorting {len(text_segments)} segments by sequence_id")
            text_segments.sort(key=lambda x: x.get('sequence_id', 0))
            logger.debug(f"[GetTranslationTexts] Sorting completed - First segment ID: {text_segments[0].get('sequence_id', 'N/A') if text_segments else 'N/A'}")
            
            # 智能过滤有内容的segments，建立映射关系
            logger.debug(f"[GetTranslationTexts] Intelligent filtering segments with content")
            content_segments = []  # 有内容的segments文本
            segment_mapping = {}   # segment索引 -> XML ID映射
            empty_count = 0
            space_count = 0
            
            for i, segment in enumerate(text_segments):
                # 兼容新旧数据格式
                space_info = segment.get('space_info', {})
                if space_info:
                    # 新格式：使用space_info判断
                    has_content = space_info.get('has_content', True)
                    is_pure_space = space_info.get('is_pure_space', False)
                    original_text = segment.get('original_text', '').strip()
                else:
                    # 旧格式：直接检查文本内容
                    original_text = segment.get('original_text', '').strip()
                    has_content = bool(original_text)
                    is_pure_space = not has_content
                
                if has_content and not is_pure_space:
                    # 有实际内容的segment，加入翻译列表
                    xml_id = f"{len(content_segments) + 1:03d}"
                    segment_mapping[i] = xml_id
                    content_segments.append(original_text)
                    logger.debug(f"[GetTranslationTexts] Content segment {xml_id}: {original_text[:50]}...") if len(content_segments) <= 5 else None
                elif is_pure_space:
                    space_count += 1
                    logger.debug(f"[GetTranslationTexts] Pure space segment at index {i}") if space_count <= 3 else None
                else:
                    empty_count += 1
            
            logger.info(f"[GetTranslationTexts] Filtering completed - Content: {len(content_segments)}, Spaces: {space_count}, Empty: {empty_count}, Total: {len(text_segments)}")
            logger.debug(f"[GetTranslationTexts] Mapping created - {len(segment_mapping)} content segments mapped")
            
            # 将映射关系存储到会话存储中，供update_translations使用
            mapping_key = f"{file_id}_mapping"
            mapping_json = json.dumps(segment_mapping, ensure_ascii=False)
            self.session.storage.set(mapping_key, mapping_json.encode('utf-8'))
            logger.debug(f"[GetTranslationTexts] Segment mapping stored with key: {mapping_key}")
            
            # Create XML segments for LLM processing
            logger.debug(f"[GetTranslationTexts] Creating XML segments for {len(content_segments)} texts")
            xml_segments = []
            for idx, text in enumerate(content_segments):
                # XML转义处理特殊字符
                escaped_text = self._xml_escape(text)
                segment_id = f"{idx+1:03d}"
                xml_segments.append(f'<segment id="{segment_id}">{escaped_text}</segment>')
            
            output_text = '\n'.join(xml_segments)
            output_length = len(output_text)
            logger.info(f"[GetTranslationTexts] XML output created - {len(xml_segments)} segments, Total length: {output_length} characters")
            
            if not output_text:
                logger.warning("[GetTranslationTexts] No translatable text found - all segments were empty")
                yield self.create_text_message("Warning: No translatable text found in the document.")
                return
            
            # Create preview (first 200 characters)
            preview = output_text[:200] + "..." if len(output_text) > 200 else output_text
            logger.debug(f"[GetTranslationTexts] Preview created - Length: {len(preview)} characters")
            logger.debug(f"[GetTranslationTexts] Preview content: {preview[:100]}...")
            
            logger.info(f"[GetTranslationTexts] Successfully retrieved {len(content_segments)} text segments for translation")
            logger.debug(f"[GetTranslationTexts] Output statistics - Total chars: {output_length}, Lines: {len(content_segments)}, Preview length: {len(preview)}")
            yield self.create_text_message(f"Retrieved {len(content_segments)} text segments for translation")
            
            # Output the texts for LLM translation
            logger.info(f"[GetTranslationTexts] Creating variable output for LLM translation")
            yield self.create_variable_message("original_texts", output_text)
            
            # Return detailed results
            result = {
                "success": True,
                "file_id": file_id,
                "original_texts": output_text,
                "text_count": len(content_segments),
                "preview": preview,
                "file_type": metadata.get("file_type", "unknown"),
                "space_count": space_count,
                "empty_count": empty_count,
                "total_segments": len(text_segments),
                "message": f"Successfully retrieved {len(content_segments)} text segments for translation"
            }
            
            yield self.create_json_message(result)
            
        except Exception as e:
            error_msg = f"Failed to retrieve translation texts: {str(e)}"
            logger.error(f"[GetTranslationTexts] Exception occurred: {error_msg}")
            yield self.create_text_message(f"Error: {error_msg}")
            yield self.create_json_message({
                "success": False,
                "file_id": tool_parameters.get("file_id", ""),
                "error": error_msg
            })
    
    def _get_metadata(self, file_id: str) -> dict:
        """Get metadata from persistent storage."""
        logger.debug(f"[GetTranslationTexts] Reading metadata for file_id: {file_id}")
        try:
            metadata_key = f"{file_id}_metadata"
            logger.debug(f"[GetTranslationTexts] Fetching metadata with key: {metadata_key}")
            metadata_data = self.session.storage.get(metadata_key)
            if not metadata_data:
                logger.warning(f"[GetTranslationTexts] No metadata found for key: {metadata_key}")
                return {}
            
            metadata = json.loads(metadata_data.decode('utf-8'))
            logger.debug(f"[GetTranslationTexts] Metadata loaded successfully - Keys: {list(metadata.keys())}")
            return metadata
        except Exception as e:
            logger.error(f"[GetTranslationTexts] Error reading metadata: {str(e)}")
            return {}
    
    def _get_text_segments(self, file_id: str) -> list:
        """Get text segments from persistent storage."""
        logger.debug(f"[GetTranslationTexts] Reading text segments for file_id: {file_id}")
        try:
            texts_key = f"{file_id}_texts"
            logger.debug(f"[GetTranslationTexts] Fetching text segments with key: {texts_key}")
            texts_data = self.session.storage.get(texts_key)
            if not texts_data:
                logger.warning(f"[GetTranslationTexts] No text segments found for key: {texts_key}")
                return []
            
            segments = json.loads(texts_data.decode('utf-8'))
            logger.debug(f"[GetTranslationTexts] Text segments loaded successfully - Count: {len(segments)}")
            
            # Log some sample segment info for debugging
            if segments:
                first_segment = segments[0]
                logger.debug(f"[GetTranslationTexts] Sample segment keys: {list(first_segment.keys())}")
                logger.debug(f"[GetTranslationTexts] Sample segment preview: {first_segment.get('original_text', '')[:50]}...")
            
            return segments
        except Exception as e:
            logger.error(f"[GetTranslationTexts] Error reading text segments: {str(e)}")
            return []
    
    def _xml_escape(self, text: str) -> str:
        """转义XML特殊字符，防止XML解析错误。"""
        if not text:
            return text
        
        try:
            # XML必须转义的字符
            escaped = text.replace('&', '&amp;')  # 必须首先转义&
            escaped = escaped.replace('<', '&lt;')
            escaped = escaped.replace('>', '&gt;')
            escaped = escaped.replace('"', '&quot;')
            escaped = escaped.replace("'", '&apos;')
            
            logger.debug(f"[GetTranslationTexts] XML escaped text: {text[:50]}... -> {escaped[:50]}...") if len(text) > 50 else None
            return escaped
        except Exception as e:
            logger.error(f"[GetTranslationTexts] Error XML escaping text: {str(e)}")
            return text  # Return original if escaping fails