import re

def extract_points(text):
    """
    Extracts the integer score inside parentheses from option texts.
    Example: 'Excelente (20 puntos)' -> 20
             'Bueno (12 puntos)' -> 12
             'Regular (4 puntos)' -> 4
             20 -> 20
    """
    if text is None:
        return 0
    if isinstance(text, (int, float)):
        return text
    
    text_str = str(text).strip()
    # Regex search for digits inside parentheses, e.g. (20 puntos) or (20)
    match = re.search(r"\((\d+)\s*puntos?\)", text_str, re.IGNORECASE)
    if match:
        return int(match.group(1))
    
    # Fallback to look for any digits in parentheses
    match_fallback = re.search(r"\((\d+)\)", text_str)
    if match_fallback:
        return int(match_fallback.group(1))
        
    return 0
