import unicodedata
import re

def normalize_name(name):
    """
    Normalizes a name string by converting it to uppercase, stripping
    leading/trailing spaces, collapsing multiple spaces, and removing accents/diacritics.
    Example: 'Solano De La Sala Aldas Paúl Fernando' -> 'SOLANO DE LA SALA ALDAS PAUL FERNANDO'
             'Jácome Macías Anne' -> 'JACOME MACIAS ANNE'
    """
    if name is None:
        return ""
    
    # Standardize to string
    name_str = str(name).strip()
    
    # Remove diacritics (accents, tildes)
    # NFKD decomposes characters (e.g., 'á' -> 'a' + 'combining acute accent')
    nfkd_form = unicodedata.normalize('NFKD', name_str)
    # Encode to ascii ignoring diacritics, then decode back
    ascii_str = nfkd_form.encode('ASCII', 'ignore').decode('ASCII')
    
    # Convert to uppercase
    upper_str = ascii_str.upper()
    
    # Collapse multiple whitespaces and remove non-word/non-space chars (except letters)
    cleaned_str = re.sub(r'\s+', ' ', upper_str)
    
    # Just in case, clean any trailing dots or punctuation
    cleaned_str = cleaned_str.strip()
    
    return cleaned_str

def split_group_names(group_str):
    """
    Splits a comma-separated string of group student names into a list of normalized names.
    Example: 'Jácome Macías Anne, Tulcán Jaya Iván' -> ['JACOME MACIAS ANNE', 'TULCAN JAYA IVAN']
    """
    if not group_str:
        return []
    
    # Split by comma
    parts = str(group_str).split(",")
    normalized_parts = []
    for part in parts:
        normalized = normalize_name(part)
        if normalized:
            normalized_parts.append(normalized)
            
    return normalized_parts
