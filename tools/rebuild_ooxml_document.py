import json
import gzip
import base64
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
from dify_plugin.file.file import File
from utils.ooxml_rebuilder import OOXMLRebuilder

# Initialize logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
logger.addHandler(plugin_logger_handler)

class RebuildOoxmlDocumentTool(Tool):
    """
    Rebuild the OOXML document with translated text while preserving all formatting.
    
    This tool reconstructs the original OOXML document by precisely replacing text elements
    with translations while maintaining all formatting, styles, and layout.
    """
    
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage, None, None]:
        logger.info("[RebuildOoxmlDocument] Starting document rebuild process")
        try:
            # Get parameters
            file_id = tool_parameters.get("file_id", "")
            input_file: File = tool_parameters.get("input_file")
            
            # Validate required parameters
            if not file_id:
                logger.error("[RebuildOoxmlDocument] Missing required parameter: file_id")
                yield self.create_text_message("Error: file_id is required.")
                return
            
            logger.info(f"[RebuildOoxmlDocument] Processing file_id: {file_id}")
            logger.debug(f"[RebuildOoxmlDocument] Starting document rebuild process")
            yield self.create_text_message(f"Starting document rebuild for file_id: {file_id}")
            
            # Read metadata
            logger.info(f"[RebuildOoxmlDocument] Reading metadata")
            metadata = self._get_metadata(file_id)
            if not metadata:
                logger.error("[RebuildOoxmlDocument] No metadata found")
                yield self.create_text_message("Error: No metadata found for this file_id. Please run extract_ooxml_text first.")
                return
            
            file_type = metadata.get("file_type", "")
            original_filename = metadata.get("original_filename", f"document.{file_type}")
            original_size_mb = metadata.get("file_size_mb", 0)
            logger.info(f"[RebuildOoxmlDocument] Metadata loaded - Type: {file_type}, Original size: {original_size_mb}MB, Filename: {original_filename}")
            
            # Read text segments with translations
            logger.info(f"[RebuildOoxmlDocument] Reading text segments")
            text_segments = self._get_text_segments(file_id)
            if not text_segments:
                logger.error("[RebuildOoxmlDocument] No text segments found")
                yield self.create_text_message("Error: No text segments found for this file_id.")
                return
            logger.info(f"[RebuildOoxmlDocument] Loaded {len(text_segments)} text segments from storage")
            
            # Check if translations are available
            translated_segments = [s for s in text_segments if s.get('translated_text', '').strip()]
            untranslated_segments = [s for s in text_segments if not s.get('translated_text', '').strip()]
            if not translated_segments:
                logger.error("[RebuildOoxmlDocument] No translations found")
                yield self.create_text_message("Error: No translations found. Please run update_translations first.")
                return
            logger.info(f"[RebuildOoxmlDocument] Translation status - Translated: {len(translated_segments)}, Untranslated: {len(untranslated_segments)}, Total: {len(text_segments)}")
            
            # Log some sample translations for debugging
            for i, segment in enumerate(translated_segments[:3]):
                original = segment.get('original_text', '')[:50]
                translated = segment.get('translated_text', '')[:50]
                logger.debug(f"[RebuildOoxmlDocument] Sample translation {i+1}: '{original}...' → '{translated}...'")
            
            # Read original file (optimized - from input_file parameter first)
            logger.info(f"[RebuildOoxmlDocument] Reading original file data")
            original_file_data = self._get_original_file_optimized(file_id, input_file)
            if not original_file_data:
                logger.error("[RebuildOoxmlDocument] Original file data not found")
                yield self.create_text_message("Error: Original file data not found. Please provide the input_file parameter.")
                return
            logger.info(f"[RebuildOoxmlDocument] Original file loaded - Size: {len(original_file_data)} bytes ({len(original_file_data)/1024/1024:.2f}MB)")
            
            file_type = metadata.get("file_type", "")
            if not file_type:
                logger.error("[RebuildOoxmlDocument] File type not found in metadata")
                yield self.create_text_message("Error: File type not found in metadata.")
                return
            logger.debug(f"[RebuildOoxmlDocument] File type confirmed: {file_type}")
            
            logger.info(f"[RebuildOoxmlDocument] Starting rebuild for {file_type.upper()} with {len(translated_segments)} translations")
            logger.debug(f"[RebuildOoxmlDocument] Rebuild parameters - File type: {file_type}, Segments: {len(text_segments)}, Translated: {len(translated_segments)}")
            yield self.create_text_message(f"Rebuilding {file_type.upper()} document with {len(translated_segments)} translations")
            
            # Initialize rebuilder
            logger.debug(f"[RebuildOoxmlDocument] Initializing OOXML rebuilder")
            rebuilder = OOXMLRebuilder()
            logger.debug(f"[RebuildOoxmlDocument] Rebuilder initialized successfully")
            
            # Rebuild the document
            rebuild_start_time = self._get_current_timestamp()
            logger.info(f"[RebuildOoxmlDocument] Starting document reconstruction - Start time: {rebuild_start_time}")
            new_file_data, replaced_count = rebuilder.rebuild_document(
                original_file_data, text_segments, file_type
            )
            rebuild_end_time = self._get_current_timestamp()
            logger.info(f"[RebuildOoxmlDocument] Document rebuilt successfully - End time: {rebuild_end_time}, Replacements: {replaced_count}")
            
            # Generate output filename
            original_filename = metadata.get("original_filename", f"document.{file_type}")
            output_filename = self._generate_translated_filename(original_filename)
            logger.debug(f"[RebuildOoxmlDocument] Filename generation - Original: {original_filename}, Output: {output_filename}")
            
            # Calculate file size
            file_size_mb = round(len(new_file_data) / (1024 * 1024), 2)
            original_size_mb = round(len(original_file_data) / (1024 * 1024), 2)
            size_change = file_size_mb - original_size_mb
            logger.info(f"[RebuildOoxmlDocument] File size comparison - Original: {original_size_mb}MB, New: {file_size_mb}MB, Change: {size_change:+.2f}MB")
            
            # Encode file as base64 for download
            logger.debug(f"[RebuildOoxmlDocument] Encoding file as base64")
            base64_content = base64.b64encode(new_file_data).decode('utf-8')
            base64_size = len(base64_content)
            download_url = f"data:application/vnd.openxmlformats-officedocument.{self._get_mime_type(file_type)};base64,{base64_content}"
            logger.info(f"[RebuildOoxmlDocument] Base64 encoding completed - Base64 size: {base64_size} characters")
            
            logger.info(f"[RebuildOoxmlDocument] Rebuild completed successfully - Replacements: {replaced_count}, File size: {file_size_mb}MB, Filename: {output_filename}")
            yield self.create_text_message(f"Successfully rebuilt document with {replaced_count} text replacements")
            
            # Create downloadable file
            logger.debug(f"[RebuildOoxmlDocument] Creating downloadable file blob")
            yield self.create_blob_message(
                new_file_data,
                meta={
                    "mime_type": f"application/vnd.openxmlformats-officedocument.{self._get_mime_type(file_type)}",
                    "filename": output_filename
                }
            )
            logger.debug(f"[RebuildOoxmlDocument] Downloadable file created successfully")
            
            # Return detailed results
            result = {
                "success": True,
                "file_id": file_id,
                "output_filename": output_filename,
                "file_size_mb": file_size_mb,
                "replaced_count": replaced_count,
                "download_url": download_url,
                "message": f"Successfully rebuilt {file_type.upper()} document with {replaced_count} text replacements"
            }
            
            yield self.create_json_message(result)
            yield self.create_variable_message("rebuild_result", json.dumps(result))
            yield self.create_variable_message("translated_document", download_url)
            
        except Exception as e:
            error_msg = f"Failed to rebuild document: {str(e)}"
            logger.error(f"[RebuildOoxmlDocument] Exception occurred: {error_msg}")
            yield self.create_text_message(f"Error: {error_msg}")
            yield self.create_json_message({
                "success": False,
                "file_id": tool_parameters.get("file_id", ""),
                "error": error_msg
            })
    
    def _get_metadata(self, file_id: str) -> dict:
        """Get metadata from persistent storage with compression support."""
        logger.debug(f"[RebuildOoxmlDocument] Reading metadata for file_id: {file_id}")
        try:
            metadata_key = f"{file_id}_metadata"
            logger.debug(f"[RebuildOoxmlDocument] Fetching metadata with key: {metadata_key}")
            metadata_data = self.session.storage.get(metadata_key)
            if not metadata_data:
                logger.warning(f"[RebuildOoxmlDocument] No metadata found for key: {metadata_key}")
                return {}
            
            # Try to decompress first (for new format)
            try:
                decompressed_data = gzip.decompress(metadata_data)
                metadata = self._fast_json_decode(decompressed_data)
                logger.debug(f"[RebuildOoxmlDocument] Compressed metadata loaded successfully - Keys: {list(metadata.keys())}")
            except (gzip.BadGzipFile, OSError):
                # Fallback to old format (uncompressed)
                metadata = json.loads(metadata_data.decode('utf-8'))
                logger.debug(f"[RebuildOoxmlDocument] Legacy metadata loaded successfully - Keys: {list(metadata.keys())}")
            
            return metadata
        except Exception as e:
            logger.error(f"[RebuildOoxmlDocument] Error reading metadata: {str(e)}")
            return {}
    
    def _get_text_segments(self, file_id: str) -> list:
        """Get text segments from persistent storage with batch and compression support."""
        logger.debug(f"[RebuildOoxmlDocument] Reading text segments for file_id: {file_id}")
        try:
            # First, check if we have batched storage
            batch_metadata = self._get_batch_metadata(file_id)
            if batch_metadata:
                return self._get_segments_batched(file_id, batch_metadata)
            
            # Try single storage (both compressed and legacy)
            texts_key = f"{file_id}_texts"
            logger.debug(f"[RebuildOoxmlDocument] Fetching text segments with key: {texts_key}")
            texts_data = self.session.storage.get(texts_key)
            if not texts_data:
                logger.warning(f"[RebuildOoxmlDocument] No text segments found for key: {texts_key}")
                return []
            
            # Try compressed format first, then legacy
            try:
                decompressed_data = gzip.decompress(texts_data)
                segments = self._fast_json_decode(decompressed_data)
                logger.debug(f"[RebuildOoxmlDocument] Compressed text segments loaded successfully - Count: {len(segments)}")
            except (gzip.BadGzipFile, OSError):
                # Fallback to old format (uncompressed)
                segments = json.loads(texts_data.decode('utf-8'))
                logger.debug(f"[RebuildOoxmlDocument] Legacy text segments loaded successfully - Count: {len(segments)}")
            
            # Restore namespace maps if optimized format is detected
            segments = self._restore_optimized_data(segments)
            
            # Count translated vs untranslated for debugging
            translated_count = sum(1 for seg in segments if seg.get('translated_text', '').strip())
            logger.debug(f"[RebuildOoxmlDocument] Translation status - Translated: {translated_count}/{len(segments)}")
            
            return segments
        except Exception as e:
            logger.error(f"[RebuildOoxmlDocument] Error reading text segments: {str(e)}")
            return []
    
    def _get_original_file_optimized(self, file_id: str, input_file: File) -> bytes:
        """Get original file data - optimized to use input_file parameter first, with KV fallback."""
        logger.debug(f"[RebuildOoxmlDocument] Getting original file data for file_id: {file_id}")
        
        # Method 1 (Preferred): Get from input_file parameter
        if input_file and hasattr(input_file, 'blob'):
            try:
                file_data = input_file.blob
                if file_data and len(file_data) > 0:
                    logger.info(f"[RebuildOoxmlDocument] Using input_file parameter - Size: {len(file_data)} bytes (performance optimized)")
                    return file_data
                else:
                    logger.warning("[RebuildOoxmlDocument] Input file parameter is empty")
            except Exception as e:
                logger.warning(f"[RebuildOoxmlDocument] Failed to get data from input_file parameter: {str(e)}")
        
        # Method 2 (Fallback): Get from KV storage (for backward compatibility)
        logger.info("[RebuildOoxmlDocument] Falling back to KV storage for original file data")
        return self._get_original_file(file_id)
    
    def _get_original_file(self, file_id: str) -> bytes:
        """Get original file data from persistent storage (legacy fallback)."""
        logger.debug(f"[RebuildOoxmlDocument] Reading original file from KV storage for file_id: {file_id}")
        try:
            file_key = f"{file_id}_original_file"
            logger.debug(f"[RebuildOoxmlDocument] Fetching original file with key: {file_key}")
            file_data = self.session.storage.get(file_key)
            if file_data:
                logger.debug(f"[RebuildOoxmlDocument] Original file loaded from KV storage - Size: {len(file_data)} bytes")
            else:
                logger.warning(f"[RebuildOoxmlDocument] No original file found in KV storage for key: {file_key}")
            return file_data
        except Exception as e:
            logger.error(f"[RebuildOoxmlDocument] Error reading original file from KV storage: {str(e)}")
            return None
    
    def _generate_translated_filename(self, original_filename: str) -> str:
        """Generate filename for translated document."""
        logger.debug(f"[RebuildOoxmlDocument] Generating translated filename from: {original_filename}")
        if '.' in original_filename:
            name, ext = original_filename.rsplit('.', 1)
            translated_filename = f"{name}_translated.{ext}"
        else:
            translated_filename = f"{original_filename}_translated"
        logger.debug(f"[RebuildOoxmlDocument] Generated filename: {translated_filename}")
        return translated_filename    
    
    def _get_current_timestamp(self) -> str:
        """Get current timestamp in ISO format."""
        from datetime import datetime
        return datetime.now().isoformat()
    
    def _get_mime_type(self, file_type: str) -> str:
        """Get MIME type suffix for file type."""
        mime_types = {
            "docx": "wordprocessingml.document",
            "xlsx": "spreadsheetml.sheet", 
            "pptx": "presentationml.presentation"
        }
        return mime_types.get(file_type, "unknown")
    
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
            
            # Decompress and decode batch metadata
            try:
                decompressed_data = gzip.decompress(batch_meta_data)
                return self._fast_json_decode(decompressed_data)
            except (gzip.BadGzipFile, OSError):
                return self._fast_json_decode(batch_meta_data)
        except Exception as e:
            logger.debug(f"[RebuildOoxmlDocument] No batch metadata found: {str(e)}")
            return {}
    
    def _get_segments_batched(self, file_id: str, batch_metadata: dict) -> list:
        """Retrieve text segments from batched storage format."""
        batch_count = batch_metadata.get('batch_count', 0)
        total_segments = batch_metadata.get('total_segments', 0)
        
        logger.debug(f"[RebuildOoxmlDocument] Loading batched segments: {batch_count} batches, {total_segments} total segments")
        
        all_segments = []
        for batch_index in range(batch_count):
            batch_key = f"{file_id}_texts_batch_{batch_index}"
            batch_data = self.session.storage.get(batch_key)
            
            if batch_data:
                try:
                    # Decompress and decode batch
                    decompressed_data = gzip.decompress(batch_data)
                    batch_segments = self._fast_json_decode(decompressed_data)
                    all_segments.extend(batch_segments)
                    logger.debug(f"[RebuildOoxmlDocument] Loaded batch {batch_index}: {len(batch_segments)} segments")
                except Exception as e:
                    logger.warning(f"[RebuildOoxmlDocument] Failed to load batch {batch_index}: {str(e)}")
            else:
                logger.warning(f"[RebuildOoxmlDocument] Batch {batch_index} not found")
        
        logger.info(f"[RebuildOoxmlDocument] Batched loading complete: {len(all_segments)}/{total_segments} segments loaded")
        
        # Restore namespace maps if optimized format is detected
        return self._restore_optimized_data(all_segments)
    
    def _restore_optimized_data(self, segments: list) -> list:
        """Restore namespace maps from optimized format using metadata registry."""
        if not segments:
            return segments
        
        # Check if any segment has namespace_ref (indicating optimized format)
        has_optimization = any(
            seg.get('xml_location', {}).get('namespace_ref') is not None
            for seg in segments
        )
        
        if not has_optimization:
            logger.debug("[RebuildOoxmlDocument] No optimization detected, segments unchanged")
            return segments
        
        # Get namespace registry from metadata
        file_id = segments[0].get('text_id', '').split('_')[0] if segments else ''
        metadata = self._get_metadata(file_id) if file_id else {}
        namespace_registry = metadata.get('namespace_registry', {})
        
        if not namespace_registry:
            logger.warning("[RebuildOoxmlDocument] Optimized format detected but no namespace registry found")
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
        
        logger.debug(f"[RebuildOoxmlDocument] Data restoration complete: {len(namespace_registry)} namespace patterns restored")
        return restored_segments