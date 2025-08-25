from typing import List, Dict
from docx import Document
from io import BytesIO


def title_key_fetcher(
    mappings: Dict[str, str], file_bytes: bytes
) -> List[Dict[str, str]]:
    """
    Given a mapping of 'Titel' to 'Nøgle' and a bytes file (Word document),
    returns a list of dictionaries with 'originalText' and 'replacementText'
    for each 'Titel' found in the document.
    """
    doc = Document(BytesIO(file_bytes))
    text = "\n".join([para.text for para in doc.paragraphs])
    text_lower = text.lower()
    # Sort titles by length descending to prioritize longer matches
    sorted_titles = sorted(mappings.keys(), key=lambda t: len(t), reverse=True)
    used_spans = []
    result = []
    for titel in sorted_titles:
        titel_lower = titel.lower()
        start = text_lower.find(titel_lower)
        if start != -1:
            end = start + len(titel_lower)
            # Check for overlap with already used spans
            overlap = False
            for s, e in used_spans:
                if not (end <= s or start >= e):
                    overlap = True
                    break
            if not overlap:
                result.append(
                    {"originalText": titel, "replacementText": mappings[titel]}
                )
                used_spans.append((start, end))
    return result
