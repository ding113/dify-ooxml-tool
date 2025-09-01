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
            
            # 简化过滤逻辑：第一步已经过滤了纯空字符串，这里直接处理所有segments
            logger.debug(f"[GetTranslationTexts] Processing all segments for translation")
            content_segments = []  # 有内容的segments文本
            segment_mapping = {}   # segment索引 -> XML ID映射
            
            for i, segment in enumerate(text_segments):
                original_text = segment.get('original_text', '')
                xml_id = f"{i + 1:03d}"
                segment_mapping[i] = xml_id
                content_segments.append(original_text)  # 保留原始文本，包括前后空格
                logger.debug(f"[GetTranslationTexts] Segment {xml_id}: {original_text[:50]}...") if len(content_segments) <= 5 else None
            
            logger.info(f"[GetTranslationTexts] Processing completed - Content segments: {len(content_segments)}, Total segments: {len(text_segments)}")
            logger.debug(f"[GetTranslationTexts] Mapping created - {len(segment_mapping)} segments mapped")
            
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
                chunks = self._create_chunks_with_overlap(xml_segments, chunk_size, max_segments_per_chunk)
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
        """Get metadata from persistent storage with simplified logic."""
        logger.debug(f"[GetTranslationTexts] Reading metadata for file_id: {file_id}")
        try:
            metadata_key = f"{file_id}_metadata"
            logger.debug(f"[GetTranslationTexts] Fetching metadata with key: {metadata_key}")
            metadata_data = self.session.storage.get(metadata_key)
            if not metadata_data:
                logger.warning(f"[GetTranslationTexts] No metadata found for key: {metadata_key}")
                return {}
            
            # Parse JSON metadata directly
            if isinstance(metadata_data, bytes):
                metadata = json.loads(metadata_data.decode('utf-8'))
            else:
                metadata = json.loads(metadata_data)
            logger.debug(f"[GetTranslationTexts] Metadata loaded successfully - Keys: {list(metadata.keys())}")
            
            return metadata
        except Exception as e:
            logger.error(f"[GetTranslationTexts] Error reading metadata: {str(e)}")
            return {}
    
    def _get_text_segments(self, file_id: str) -> list:
        """Get text segments from persistent storage with simplified logic."""
        logger.debug(f"[GetTranslationTexts] Reading text segments for file_id: {file_id}")
        try:
            texts_key = f"{file_id}_texts"
            logger.debug(f"[GetTranslationTexts] Fetching text segments with key: {texts_key}")
            texts_data = self.session.storage.get(texts_key)
            if not texts_data:
                logger.warning(f"[GetTranslationTexts] No text segments found for key: {texts_key}")
                return []
            
            # Parse JSON segments directly
            if isinstance(texts_data, bytes):
                segments = json.loads(texts_data.decode('utf-8'))
            else:
                segments = json.loads(texts_data)
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
    
    def _create_chunks_with_overlap(self, xml_segments: list, chunk_size: int, max_segments_per_chunk: int) -> list:
        """
        创建带5%重叠的XML段落块，解决边界漏译问题。
        
        Args:
            xml_segments: XML格式的文本段落列表
            chunk_size: 每个块的最大字符数
            max_segments_per_chunk: 每个块的最大段落数
            
        Returns:
            分块后的XML字符串列表，相邻分块间有5%重叠
        """
        if not xml_segments:
            return []
        
        logger.debug(f"[GetTranslationTexts] Starting overlap chunking - {len(xml_segments)} segments, chunk_size: {chunk_size}, max_segments: {max_segments_per_chunk}")
        
        chunks = []
        overlap_size = max(1, int(max_segments_per_chunk * 0.05))  # 5%重叠
        start_index = 0
        
        while start_index < len(xml_segments):
            current_chunk = []
            current_chunk_size = 0
            
            # 计算当前块的段落范围
            end_index = min(start_index + max_segments_per_chunk, len(xml_segments))
            
            for i in range(start_index, end_index):
                segment = xml_segments[i]
                segment_size = len(segment)
                
                # 检查是否超过字符限制
                if current_chunk and current_chunk_size + segment_size + 1 > chunk_size:
                    break
                
                current_chunk.append(segment)
                current_chunk_size += segment_size + 1  # +1 for newline
            
            if current_chunk:
                chunk_text = '\n'.join(current_chunk)
                chunks.append(chunk_text)
                logger.debug(f"[GetTranslationTexts] Overlap chunk {len(chunks)} created - {len(current_chunk)} segments, {current_chunk_size} chars")
            
            # 计算下一块的起始位置（有重叠）
            if start_index + len(current_chunk) >= len(xml_segments):
                break
                
            # 下一块起始位置：当前块结束位置 - 重叠大小
            start_index = start_index + len(current_chunk) - overlap_size
        
        logger.info(f"[GetTranslationTexts] Overlap chunking complete - Created {len(chunks)} chunks with {overlap_size} segments overlap")
        
        # 记录重叠统计信息
        for i, chunk in enumerate(chunks):
            chunk_lines = chunk.count('\n') + 1
            logger.debug(f"[GetTranslationTexts] Overlap chunk {i+1}: {len(chunk)} chars, {chunk_lines} segments")
        
        return chunks