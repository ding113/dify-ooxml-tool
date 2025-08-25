import json
import gzip
import logging
from collections.abc import Generator
from typing import Any
import concurrent.futures
import threading

# Use orjson for high-performance JSON serialization, fallback to standard json
try:
    import orjson
    ORJSON_AVAILABLE = True
except ImportError:
    ORJSON_AVAILABLE = False

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.config.logger_format import plugin_logger_handler
from dify_plugin.file.file import File, FileType
from utils.ooxml_parser import OOXMLParser

# Initialize logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
logger.addHandler(plugin_logger_handler)

class ExtractOoxmlTextTool(Tool):
    """
    Extract translatable text from OOXML documents (DOCX, XLSX, PPTX).
    
    This tool parses OOXML files and extracts all translatable text while preserving
    exact XML locations for format-preserving translation reconstruction.
    """
    
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage, None, None]:
        logger.info("[ExtractOoxmlText] Starting OOXML text extraction")
        try:
            # Get parameters
            input_file: File = tool_parameters.get("input_file")
            file_id = tool_parameters.get("file_id", "")
            
            # Validate required parameters
            if not file_id:
                logger.error("[ExtractOoxmlText] Missing required parameter: file_id")
                yield self.create_text_message("Error: file_id is required.")
                return
            
            if not input_file or not isinstance(input_file, File):
                logger.error("[ExtractOoxmlText] Missing or invalid file parameter: input_file must be a valid File object")
                yield self.create_text_message("Error: A valid OOXML file (DOCX, XLSX, or PPTX) must be uploaded.")
                return
            
            logger.info(f"[ExtractOoxmlText] Processing file_id: {file_id}")
            logger.debug(f"[ExtractOoxmlText] Input file: {input_file.filename}, size: {len(input_file.blob)} bytes")
            yield self.create_text_message(f"Starting OOXML text extraction for file_id: {file_id}")
            
            # Initialize parser
            parser = OOXMLParser()
            
            # Get file data
            logger.info(f"[ExtractOoxmlText] Retrieving file data from uploaded file: {input_file.filename}")
            file_data, file_size = self._get_file_data(input_file)
            if not file_data:
                logger.error("[ExtractOoxmlText] Failed to retrieve file data")
                yield self.create_text_message("Error: Failed to retrieve file data.")
                return
            logger.info(f"[ExtractOoxmlText] File data retrieved successfully - Size: {file_size} bytes ({file_size/1024/1024:.2f} MB)")
            
            # Detect file type
            logger.info("[ExtractOoxmlText] Detecting file type")
            file_type = parser.detect_file_type(file_data)
            if not file_type:
                logger.error("[ExtractOoxmlText] Invalid file type detected - not a valid OOXML file")
                yield self.create_text_message("Error: Not a valid OOXML file (DOCX, XLSX, PPTX).")
                return
            logger.info(f"[ExtractOoxmlText] File type detected successfully: {file_type.upper()}")
            
            yield self.create_text_message(f"Detected file type: {file_type.upper()}")
            
            # Parse the file
            logger.info(f"[ExtractOoxmlText] Starting {file_type.upper()} parsing")
            parse_start_time = self._get_current_timestamp()
            parse_result = parser.parse_file(file_data, file_type)
            parse_end_time = self._get_current_timestamp()
            logger.debug(f"[ExtractOoxmlText] Parsing completed - Start: {parse_start_time}, End: {parse_end_time}")
            
            if not parse_result.get("success", False):
                error_msg = parse_result.get("error", "Unknown parsing error")
                logger.error(f"[ExtractOoxmlText] Parsing failed: {error_msg}")
                yield self.create_text_message(f"Error: {error_msg}")
                return
            
            text_segments = parse_result.get("text_segments", [])
            supported_elements = parse_result.get("supported_elements", [])
            logger.info(f"[ExtractOoxmlText] Parsing successful - Found {len(text_segments)} text segments")
            logger.debug(f"[ExtractOoxmlText] Supported elements: {supported_elements}")
            
            # Assign sequence IDs to ensure proper ordering
            logger.info(f"[ExtractOoxmlText] Assigning sequence IDs to {len(text_segments)} text segments")
            for i, segment in enumerate(text_segments):
                segment["sequence_id"] = i
                logger.debug(f"[ExtractOoxmlText] Segment {i}: {segment.get('original_text', '')[:50]}...") if i < 5 else None  # Log first 5 for debugging
            
            # Store data in persistent storage (optimized - without original file)
            storage_start_time = self._get_current_timestamp()
            logger.info(f"[ExtractOoxmlText] Storing extraction data to persistent storage - {len(text_segments)} segments (original file excluded for performance)")
            self._store_extraction_data(file_id, {
                "file_type": file_type,
                "original_filename": input_file.filename or f"document.{file_type}",
                "extraction_timestamp": self._get_current_timestamp(),
                "total_text_count": len(text_segments),
                "file_size_mb": round(file_size / (1024 * 1024), 2),
                "supported_elements": supported_elements
            }, text_segments)
            storage_end_time = self._get_current_timestamp()
            logger.info(f"[ExtractOoxmlText] Storage completed successfully - Start: {storage_start_time}, End: {storage_end_time}")
            
            logger.info(f"[ExtractOoxmlText] Extraction completed successfully - {len(text_segments)} text segments, {file_size/1024/1024:.2f}MB file")
            yield self.create_text_message(f"Successfully extracted {len(text_segments)} text segments")
            
            # Return lightweight summary (optimized to avoid stdout buffer overflow)
            result = {
                "success": True,
                "file_id": file_id,
                "file_type": file_type,
                "extracted_text_count": len(text_segments),
                "file_size_mb": round(file_size / (1024 * 1024), 2),
                "supported_elements": supported_elements,
                "storage_keys": {
                    "metadata": f"{file_id}_metadata",
                    "text_segments": f"{file_id}_texts", 
                    "original_file": f"{file_id}_original_file"
                },
                "message": f"Successfully extracted {len(text_segments)} text segments from {file_type.upper()} file"
            }
            
            yield self.create_json_message(result)
            # Removed large JSON variable output to prevent stdout buffer overflow
            # Data is safely stored in KV storage and accessible via storage_keys
            
        except Exception as e:
            error_msg = f"Extraction failed: {str(e)}"
            logger.error(f"[ExtractOoxmlText] Exception occurred: {error_msg}")
            yield self.create_text_message(f"Error: {error_msg}")
            yield self.create_json_message({
                "success": False,
                "file_id": tool_parameters.get("file_id", ""),
                "error": error_msg
            })
    
    def _get_file_data(self, input_file: File) -> tuple[bytes, int]:
        """Get file data from uploaded File object."""
        logger.debug(f"[ExtractOoxmlText] Getting file data from uploaded file: {input_file.filename}")
        try:
            file_data = input_file.blob
            file_size = len(file_data)
            logger.debug(f"[ExtractOoxmlText] File data retrieved successfully - Size: {file_size} bytes")
            return file_data, file_size
        except Exception as e:
            logger.error(f"[ExtractOoxmlText] Failed to get file data: {str(e)}")
            raise Exception(f"Failed to get file data: {str(e)}")
    
    def _store_extraction_data(self, file_id: str, metadata: dict, text_segments: list):
        """Store extraction data with high-performance optimizations (compression + orjson + batching)."""
        storage_start = self._get_current_timestamp()
        logger.debug(f"[ExtractOoxmlText] High-performance storage for file_id: {file_id}, segments: {len(text_segments)}")
        
        try:
            # Optimize data structure before serialization
            optimized_metadata, optimized_segments = self._optimize_storage_data(metadata, text_segments)
            
            # Use high-performance JSON serialization
            metadata_json = self._fast_json_encode(optimized_metadata)
            
            # Compress and store metadata
            metadata_compressed = gzip.compress(metadata_json)
            metadata_key = f"{file_id}_metadata"
            metadata_original_size = len(metadata_json)
            metadata_compressed_size = len(metadata_compressed)
            compression_ratio = (1 - metadata_compressed_size / metadata_original_size) * 100
            
            logger.debug(f"[ExtractOoxmlText] Metadata compression: {metadata_original_size}→{metadata_compressed_size} bytes ({compression_ratio:.1f}% reduction)")
            
            # Store text segments in optimized batches
            if len(optimized_segments) > 50:  # Use batching for large datasets
                total_segments_size = self._store_segments_batched(file_id, optimized_segments)
            else:
                # Small datasets - store directly with compression
                segments_json = self._fast_json_encode(optimized_segments)
                segments_compressed = gzip.compress(segments_json)
                segments_key = f"{file_id}_texts"
                self.session.storage.set(segments_key, segments_compressed)
                total_segments_size = len(segments_compressed)
                logger.debug(f"[ExtractOoxmlText] Segments compression: {len(segments_json)}→{len(segments_compressed)} bytes")
            
            # Store metadata and log results
            self.session.storage.set(metadata_key, metadata_compressed)
            
            total_stored_size = metadata_compressed_size + total_segments_size
            storage_end = self._get_current_timestamp()
            
            logger.info(f"[ExtractOoxmlText] High-performance storage completed - Size: {total_stored_size} bytes ({total_stored_size/1024/1024:.3f}MB)")
            logger.info(f"[ExtractOoxmlText] Storage time: {storage_start} → {storage_end}")
            
        except Exception as e:
            logger.error(f"[ExtractOoxmlText] High-performance storage failed: {str(e)}")
            logger.debug(f"[ExtractOoxmlText] Storage error details - segments_count: {len(text_segments)}")
            raise Exception(f"Failed to store extraction data: {str(e)}")
    
    def _get_current_timestamp(self) -> str:
        """Get current timestamp in ISO format."""
        from datetime import datetime
        return datetime.now().isoformat()
    
    def _fast_json_encode(self, data) -> bytes:
        """High-performance JSON encoding with orjson fallback."""
        if ORJSON_AVAILABLE:
            return orjson.dumps(data, option=orjson.OPT_NON_STR_KEYS)
        else:
            return json.dumps(data, ensure_ascii=False).encode('utf-8')
    
    def _optimize_storage_data(self, metadata: dict, text_segments: list) -> tuple[dict, list]:
        """Optimize data structures before storage to reduce redundancy."""
        # Extract common namespace mappings to reduce redundancy
        common_namespaces = {}
        namespace_id_map = {}
        namespace_counter = 0
        
        # Build common namespace registry
        for segment in text_segments:
            xml_location = segment.get('xml_location', {})
            namespace_map = xml_location.get('namespace_map', {})
            namespace_key = json.dumps(namespace_map, sort_keys=True)
            
            if namespace_key not in namespace_id_map:
                namespace_id_map[namespace_key] = f"ns_{namespace_counter}"
                common_namespaces[f"ns_{namespace_counter}"] = namespace_map
                namespace_counter += 1
        
        # Optimize segments by referencing common namespaces
        optimized_segments = []
        for segment in text_segments:
            optimized_segment = segment.copy()
            xml_location = optimized_segment.get('xml_location', {})
            namespace_map = xml_location.get('namespace_map', {})
            namespace_key = json.dumps(namespace_map, sort_keys=True)
            
            # Replace namespace_map with reference ID
            if namespace_key in namespace_id_map:
                xml_location['namespace_ref'] = namespace_id_map[namespace_key]
                xml_location.pop('namespace_map', None)  # Remove redundant data
            
            optimized_segments.append(optimized_segment)
        
        # Add namespace registry to metadata
        optimized_metadata = metadata.copy()
        optimized_metadata['namespace_registry'] = common_namespaces
        optimized_metadata['optimization_applied'] = True
        
        logger.debug(f"[ExtractOoxmlText] Data optimization: {len(common_namespaces)} unique namespace patterns identified")
        return optimized_metadata, optimized_segments
    
    def _store_segments_batched(self, file_id: str, segments: list, batch_size: int = 50) -> int:
        """Store text segments in compressed batches for optimal performance."""
        total_size = 0
        batch_count = (len(segments) + batch_size - 1) // batch_size  # Ceiling division
        
        logger.debug(f"[ExtractOoxmlText] Batched storage: {len(segments)} segments in {batch_count} batches of {batch_size}")
        
        # Use threading for parallel batch processing
        def store_batch(batch_index: int, batch_data: list) -> int:
            batch_json = self._fast_json_encode(batch_data)
            batch_compressed = gzip.compress(batch_json)
            batch_key = f"{file_id}_texts_batch_{batch_index}"
            self.session.storage.set(batch_key, batch_compressed)
            
            compression_ratio = (1 - len(batch_compressed) / len(batch_json)) * 100
            logger.debug(f"[ExtractOoxmlText] Batch {batch_index}: {len(batch_json)}→{len(batch_compressed)} bytes ({compression_ratio:.1f}% reduction)")
            return len(batch_compressed)
        
        with concurrent.futures.ThreadPoolExecutor(max_workers=3) as executor:
            futures = []
            for i in range(0, len(segments), batch_size):
                batch_index = i // batch_size
                batch_data = segments[i:i + batch_size]
                future = executor.submit(store_batch, batch_index, batch_data)
                futures.append(future)
            
            # Collect results
            for future in concurrent.futures.as_completed(futures):
                total_size += future.result()
        
        # Store batch metadata
        batch_metadata = {
            'batch_count': batch_count,
            'batch_size': batch_size,
            'total_segments': len(segments)
        }
        batch_meta_json = self._fast_json_encode(batch_metadata)
        batch_meta_compressed = gzip.compress(batch_meta_json)
        batch_meta_key = f"{file_id}_batch_metadata"
        self.session.storage.set(batch_meta_key, batch_meta_compressed)
        total_size += len(batch_meta_compressed)
        
        logger.debug(f"[ExtractOoxmlText] Batched storage completed: {batch_count} batches, {total_size} bytes total")
        return total_size