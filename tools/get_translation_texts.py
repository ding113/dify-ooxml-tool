import json
import gzip
import logging
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
            output_format = tool_parameters.get("output_format", "string")
            chunk_size = tool_parameters.get("chunk_size", 1500)
            max_segments_per_chunk = tool_parameters.get("max_segments_per_chunk", 50)
            
            # Validate required parameters
            if not file_id:
                logger.error("[GetTranslationTexts] Missing required parameter: file_id")
                yield self.create_text_message("Error: file_id is required.")
                return
            
            logger.info(f"[GetTranslationTexts] Parameters - Output format: {output_format}, Chunk size: {chunk_size}, Max segments: {max_segments_per_chunk}")
            
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
            
            if not xml_segments:
                logger.warning("[GetTranslationTexts] No translatable text found - all segments were empty")
                yield self.create_text_message("Warning: No translatable text found in the document.")
                return
            
            # Handle different output formats
            if output_format == "array":
                # Create chunks for parallel processing
                logger.info(f"[GetTranslationTexts] Creating chunks for array output format")
                chunks = self._create_chunks(xml_segments, chunk_size, max_segments_per_chunk)
                chunk_count = len(chunks)
                
                # Calculate total length for preview
                output_text = '\n'.join(xml_segments)
                output_length = len(output_text)
                preview = output_text[:200] + "..." if len(output_text) > 200 else output_text
                
                logger.info(f"[GetTranslationTexts] Created {chunk_count} chunks from {len(xml_segments)} segments")
                logger.debug(f"[GetTranslationTexts] Chunk statistics - Total chars: {output_length}, Chunks: {chunk_count}")
                
                yield self.create_text_message(f"Created {chunk_count} chunks from {len(content_segments)} text segments for parallel translation")
                
                # Output chunks as array for iteration node
                yield self.create_variable_message("chunks", chunks)
                
                # Return detailed results for array format
                result = {
                    "success": True,
                    "file_id": file_id,
                    "chunks": chunks,
                    "text_count": len(content_segments),
                    "chunk_count": chunk_count,
                    "preview": preview,
                    "file_type": metadata.get("file_type", "unknown"),
                    "space_count": space_count,
                    "empty_count": empty_count,
                    "total_segments": len(text_segments),
                    "message": f"Successfully created {chunk_count} chunks from {len(content_segments)} text segments for parallel translation"
                }
            else:
                # Single string output (original behavior)
                output_text = '\n'.join(xml_segments)
                output_length = len(output_text)
                preview = output_text[:200] + "..." if len(output_text) > 200 else output_text
                
                logger.info(f"[GetTranslationTexts] XML output created - {len(xml_segments)} segments, Total length: {output_length} characters")
                logger.debug(f"[GetTranslationTexts] Preview created - Length: {len(preview)} characters")
                
                yield self.create_text_message(f"Retrieved {len(content_segments)} text segments for translation")
                
                # Output the texts for LLM translation
                logger.info(f"[GetTranslationTexts] Creating variable output for LLM translation")
                yield self.create_variable_message("original_texts", output_text)
                
                # Return detailed results for string format
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
        """Get metadata from persistent storage with compression support."""
        logger.debug(f"[GetTranslationTexts] Reading metadata for file_id: {file_id}")
        try:
            metadata_key = f"{file_id}_metadata"
            logger.debug(f"[GetTranslationTexts] Fetching metadata with key: {metadata_key}")
            metadata_data = self.session.storage.get(metadata_key)
            if not metadata_data:
                logger.warning(f"[GetTranslationTexts] No metadata found for key: {metadata_key}")
                return {}
            
            # Try compressed format first (new format)
            try:
                decompressed_data = gzip.decompress(metadata_data)
                metadata = self._fast_json_decode(decompressed_data)
                logger.debug(f"[GetTranslationTexts] Compressed metadata loaded successfully - Keys: {list(metadata.keys())}")
            except (gzip.BadGzipFile, OSError):
                # Fallback to legacy format (uncompressed)
                metadata = json.loads(metadata_data.decode('utf-8'))
                logger.debug(f"[GetTranslationTexts] Legacy metadata loaded successfully - Keys: {list(metadata.keys())}")
            
            return metadata
        except Exception as e:
            logger.error(f"[GetTranslationTexts] Error reading metadata: {str(e)}")
            return {}
    
    def _get_text_segments(self, file_id: str) -> list:
        """Get text segments from persistent storage with compression and batch support."""
        logger.debug(f"[GetTranslationTexts] Reading text segments for file_id: {file_id}")
        try:
            # First, check if we have batched storage
            batch_metadata = self._get_batch_metadata(file_id)
            if batch_metadata:
                return self._get_segments_batched(file_id, batch_metadata)
            
            # Try single storage (both compressed and legacy)
            texts_key = f"{file_id}_texts"
            logger.debug(f"[GetTranslationTexts] Fetching text segments with key: {texts_key}")
            texts_data = self.session.storage.get(texts_key)
            if not texts_data:
                logger.warning(f"[GetTranslationTexts] No text segments found for key: {texts_key}")
                return []
            
            # Try compressed format first (new format)
            try:
                decompressed_data = gzip.decompress(texts_data)
                segments = self._fast_json_decode(decompressed_data)
                logger.debug(f"[GetTranslationTexts] Compressed text segments loaded successfully - Count: {len(segments)}")
            except (gzip.BadGzipFile, OSError):
                # Fallback to legacy format (uncompressed)
                segments = json.loads(texts_data.decode('utf-8'))
                logger.debug(f"[GetTranslationTexts] Legacy text segments loaded successfully - Count: {len(segments)}")
            
            # Restore namespace maps if optimized format is detected
            segments = self._restore_optimized_data(segments, file_id)
            
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
            logger.debug(f"[GetTranslationTexts] No batch metadata found: {str(e)}")
            return {}
    
    def _get_segments_batched(self, file_id: str, batch_metadata: dict) -> list:
        """Retrieve text segments from batched storage format."""
        batch_count = batch_metadata.get('batch_count', 0)
        total_segments = batch_metadata.get('total_segments', 0)
        
        logger.debug(f"[GetTranslationTexts] Loading batched segments: {batch_count} batches, {total_segments} total segments")
        
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
                    logger.debug(f"[GetTranslationTexts] Loaded batch {batch_index}: {len(batch_segments)} segments")
                except Exception as e:
                    logger.warning(f"[GetTranslationTexts] Failed to load batch {batch_index}: {str(e)}")
            else:
                logger.warning(f"[GetTranslationTexts] Batch {batch_index} not found")
        
        logger.info(f"[GetTranslationTexts] Batched loading complete: {len(all_segments)}/{total_segments} segments loaded")
        
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
            logger.debug("[GetTranslationTexts] No optimization detected, segments unchanged")
            return segments
        
        # Get namespace registry from metadata
        metadata = self._get_metadata(file_id)
        namespace_registry = metadata.get('namespace_registry', {})
        
        if not namespace_registry:
            logger.warning("[GetTranslationTexts] Optimized format detected but no namespace registry found")
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
        
        logger.debug(f"[GetTranslationTexts] Data restoration complete: {len(namespace_registry)} namespace patterns restored")
        return restored_segments
    
    def _create_chunks(self, xml_segments: list, chunk_size: int, max_segments_per_chunk: int) -> list:
        """
        智能创建XML段落块，优化并行翻译处理。
        
        Args:
            xml_segments: XML格式的文本段落列表
            chunk_size: 每个块的最大字符数
            max_segments_per_chunk: 每个块的最大段落数
            
        Returns:
            分块后的XML字符串列表，每个字符串包含多个XML段落
        """
        if not xml_segments:
            return []
        
        logger.debug(f"[GetTranslationTexts] Starting chunking process - {len(xml_segments)} segments, chunk_size: {chunk_size}, max_segments: {max_segments_per_chunk}")
        
        chunks = []
        current_chunk = []
        current_chunk_size = 0
        
        for segment in xml_segments:
            segment_size = len(segment)
            
            # 检查是否应该开始新的块
            should_start_new_chunk = (
                # 当前块已有内容且添加新段落会超过大小限制
                (current_chunk and current_chunk_size + segment_size + 1 > chunk_size) or  # +1 for newline
                # 或者已达到最大段落数限制
                (current_chunk and len(current_chunk) >= max_segments_per_chunk)
            )
            
            if should_start_new_chunk:
                # 完成当前块并开始新块
                if current_chunk:
                    chunk_text = '\n'.join(current_chunk)
                    chunks.append(chunk_text)
                    logger.debug(f"[GetTranslationTexts] Chunk {len(chunks)} created - {len(current_chunk)} segments, {current_chunk_size} chars")
                
                # 开始新块
                current_chunk = [segment]
                current_chunk_size = segment_size
            else:
                # 添加到当前块
                current_chunk.append(segment)
                current_chunk_size += segment_size + 1  # +1 for newline
        
        # 处理最后一个块
        if current_chunk:
            chunk_text = '\n'.join(current_chunk)
            chunks.append(chunk_text)
            logger.debug(f"[GetTranslationTexts] Final chunk {len(chunks)} created - {len(current_chunk)} segments, {current_chunk_size} chars")
        
        logger.info(f"[GetTranslationTexts] Chunking complete - Created {len(chunks)} chunks from {len(xml_segments)} segments")
        
        # 记录每个块的统计信息
        for i, chunk in enumerate(chunks):
            chunk_lines = chunk.count('\n') + 1
            logger.debug(f"[GetTranslationTexts] Chunk {i+1}: {len(chunk)} chars, {chunk_lines} segments")
        
        return chunks