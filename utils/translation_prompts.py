"""
Translation prompt templates for XML-based translation workflow.

This module provides standardized prompt templates that ensure LLM follows
XML segment structure and maintains 1:1 correspondence between original and translated texts.
"""

def get_xml_translation_prompt(source_lang: str = "Chinese", target_lang: str = "English") -> str:
    """
    Generate XML format translation system prompt.
    
    This prompt ensures LLM maintains the exact segment structure and IDs,
    preventing the common issue of merging or splitting text segments.
    
    Args:
        source_lang: Source language name (e.g., "Chinese", "Japanese")
        target_lang: Target language name (e.g., "English", "French")
        
    Returns:
        Formatted system prompt string for XML translation
    """
    return f"""Please translate the following text segments from {source_lang} to {target_lang}.

CRITICAL REQUIREMENTS:
1. Keep the EXACT same number of segments
2. Maintain the XML structure with id attributes
3. Do NOT merge, split, or skip any segments
4. Preserve the sequential ID numbers (001, 002, 003, etc.)
5. Only translate the text content inside the tags
6. Keep empty segments as empty (if any exist)
7. Preserve the exact XML format in your response

Input format: <segment id="001">Original text</segment>
Output format: <segment id="001">Translated text</segment>

Example:
Input:
<segment id="001">案例</segment>
<segment id="002">張紘賓</segment>
<segment id="003">2025/4/28</segment>

Expected Output:
<segment id="001">Case</segment>
<segment id="002">Zhang Hongbin</segment>
<segment id="003">2025/4/28</segment>

IMPORTANT NOTES:
- Numbers, dates, and proper names may not need translation
- Keep technical terms consistent
- Maintain professional translation quality
- If unsure about a term, keep it in original language

Now translate the following segments:"""

def get_xml_translation_prompt_chinese() -> str:
    """获取中文到英文的XML翻译提示模板。"""
    return get_xml_translation_prompt("Chinese", "English")

def get_xml_translation_prompt_japanese() -> str:
    """获取日文到英文的XML翻译提示模板。"""
    return get_xml_translation_prompt("Japanese", "English")

def get_batch_translation_prompt(source_lang: str = "Chinese", target_lang: str = "English", batch_size: int = 20) -> str:
    """
    Generate prompt for batch translation processing.
    
    For large documents, this prompt handles smaller batches to ensure better quality
    and reduce the chance of segment merging by LLM.
    
    Args:
        source_lang: Source language name
        target_lang: Target language name
        batch_size: Number of segments in current batch
        
    Returns:
        Batch-optimized translation prompt
    """
    return f"""Please translate this batch of {batch_size} text segments from {source_lang} to {target_lang}.

BATCH TRANSLATION REQUIREMENTS:
1. This is a batch of {batch_size} segments from a larger document
2. Maintain EXACT segment count and ID sequence
3. Do NOT merge segments even if they seem related
4. Each segment should be translated independently
5. Preserve the XML structure precisely

Format requirements:
- Input: <segment id="XXX">Original text</segment>
- Output: <segment id="XXX">Translated text</segment>

Quality guidelines:
- Professional translation quality
- Consistent terminology within the batch
- Preserve technical terms and proper names appropriately
- Maintain original formatting and punctuation style

Current batch to translate:"""

def get_quality_check_prompt() -> str:
    """
    Generate prompt for translation quality verification.
    
    This can be used as a second-pass prompt to verify translation quality
    and segment correspondence.
    """
    return """Please review the following XML translation for quality and correctness.

CHECK FOR:
1. Segment count consistency (same number of input and output segments)
2. ID sequence preservation (001, 002, 003, etc.)
3. Translation quality and accuracy
4. XML format correctness
5. No merged or split segments

If you find any issues, please provide:
1. A corrected version
2. Brief explanation of what was fixed

ORIGINAL and TRANSLATED segments to review:"""

def get_repair_prompt() -> str:
    """
    Generate prompt for repairing malformed XML translation output.
    
    When XML parsing fails, this prompt can attempt to repair common issues.
    """
    return """The following translation output has XML formatting issues. Please repair it to proper XML format.

REPAIR REQUIREMENTS:
1. Fix any broken XML tags
2. Ensure all segments have proper id attributes
3. Maintain segment count
4. Preserve translation content
5. Use proper XML escaping for special characters

Expected format:
<segment id="001">Translation content</segment>
<segment id="002">Translation content</segment>

Please repair this malformed XML:"""

# Predefined language pairs for common use cases
LANGUAGE_PAIRS = {
    "zh-en": ("Chinese", "English"),
    "ja-en": ("Japanese", "English"), 
    "ko-en": ("Korean", "English"),
    "zh-ja": ("Chinese", "Japanese"),
    "en-zh": ("English", "Chinese"),
    "en-ja": ("English", "Japanese")
}

def get_translation_prompt_by_code(lang_code: str) -> str:
    """
    Get translation prompt by language code.
    
    Args:
        lang_code: Language pair code (e.g., "zh-en", "ja-en")
        
    Returns:
        Appropriate translation prompt
        
    Raises:
        ValueError: If language code is not supported
    """
    if lang_code not in LANGUAGE_PAIRS:
        raise ValueError(f"Unsupported language pair: {lang_code}. Supported: {list(LANGUAGE_PAIRS.keys())}")
    
    source_lang, target_lang = LANGUAGE_PAIRS[lang_code]
    return get_xml_translation_prompt(source_lang, target_lang)