"""
Balaram to Unicode conversion module
Converts Balaram font characters to proper Unicode equivalents
"""

# Complete Balaram to Unicode character mapping
balaram_map = {
    # Lowercase diacritical marks
    'ä': 'ā',   # long a
    'é': 'ī',   # long i
    'ü': 'ū',   # long u
    'å': 'ṛ',   # vocalic r
    'è': 'ṝ',   # long vocalic r
    'ì': 'ṅ',   # nasal n (velar)
    'ï': 'ñ',   # nasal n (palatal)
    'ö': 'ṭ',   # retroflex t
    'ò': 'ḍ',   # retroflex d
    'ë': 'ṇ',   # retroflex n
    'ç': 'ś',   # palatal s
    'à': 'ṁ',   # anusvara
    'ù': 'ḥ',   # visarga
    'ÿ': 'ḷ',   # vocalic l
    'û': 'ḹ',   # long vocalic l
    'ý': 'ẏ',   # y with dot above
    'ñ': 'ṣ',   # retroflex s
    
    # Uppercase diacritical marks
    'Ä': 'Ā',   # long A
    'É': 'Ī',   # long I
    'Ü': 'Ū',   # long U
    'Å': 'Ṛ',   # vocalic R
    'È': 'Ṝ',   # long vocalic R
    'Ì': 'Ṅ',   # nasal N (velar)
    'Ï': 'Ñ',   # nasal N (palatal)
    'Ö': 'Ṭ',   # retroflex T
    'Ò': 'Ḍ',   # retroflex D
    'Ë': 'Ṇ',   # retroflex N
    'Ç': 'Ś',   # palatal S
    'À': 'Ṁ',   # Anusvara
    'Ù': 'Ḥ',   # Visarga
    'ß': 'Ḷ',   # vocalic L
    'Ý': 'Ẏ',   # Y with dot above
    'Ñ': 'Ṣ',   # retroflex S

    # Special characters
    '~': 'ɱ',   # special m
    "'": "'",   # straight apostrophe
    '…': '…',   # ellipsis
    '‘': "'",   # left single curly quote
    '’': "'",   # right single curly quote
    '“': '"',   # left double curly quote
    '”': '"',   # right double curly quote

}

def convert_balaram_to_unicode(text: str) -> str:
    """
    Convert Balaram font text to Unicode.
    
    Args:
        text (str): Text string containing Balaram font characters
        
    Returns:
        str: Text converted to proper Unicode characters
    """
    if not text:
        return text
        
    # Convert each character using the mapping
    return ''.join(balaram_map.get(char, char) for char in text)

def get_conversion_stats(text: str) -> dict:
    """
    Get statistics about conversions that would be made.
    
    Args:
        text (str): Input text to analyze
        
    Returns:
        dict: Statistics about conversions
    """
    if not text:
        return {'total_chars': 0, 'converted_chars': 0, 'conversion_rate': 0}
    
    total_chars = len(text)
    converted_chars = sum(1 for char in text if char in balaram_map)
    conversion_rate = (converted_chars / total_chars * 100) if total_chars > 0 else 0
    
    return {
        'total_chars': total_chars,
        'converted_chars': converted_chars,
        'conversion_rate': round(conversion_rate, 2)
    }

def preview_conversion(text: str, max_length: int = 100) -> tuple:
    """
    Preview what the conversion would look like.
    
    Args:
        text (str): Original text
        max_length (int): Maximum length to preview
        
    Returns:
        tuple: (original_preview, converted_preview, has_changes)
    """
    if not text:
        return ("", "", False)
    
    preview_text = text[:max_length]
    if len(text) > max_length:
        preview_text += "..."
    
    converted = convert_balaram_to_unicode(preview_text)
    has_changes = preview_text != converted
    
    return (preview_text, converted, has_changes)

# Export the main function and mapping for external use
__all__ = ['convert_balaram_to_unicode', 'balaram_map', 'get_conversion_stats', 'preview_conversion']
