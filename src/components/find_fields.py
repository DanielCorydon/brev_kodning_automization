from docx.shared import RGBColor


def get_run_colors(run):
    """
    Extract text color and highlight color from a run.
    Returns a tuple (text_color, highlight_color) where colors are normalized strings.
    """
    text_color = None
    highlight_color = None

    # Get text color
    if run.font.color and run.font.color.rgb:
        text_color = str(run.font.color.rgb)

    # Get highlight color
    highlights = run._element.xpath(".//w:highlight")
    for highlight in highlights:
        from docx.oxml.ns import qn

        bg_color = highlight.get(qn("w:val"))
        if bg_color:
            highlight_color = bg_color
            break

    # Also check for shading elements
    if not highlight_color:
        shadings = run._element.xpath(".//w:shd")
        for shading in shadings:
            from docx.oxml.ns import qn

            fill_color = shading.get(qn("w:fill"))
            if fill_color and fill_color != "auto":
                highlight_color = fill_color
                break

    return (text_color, highlight_color)


def find_title_match_with_mf_prefix(
    segment_text: str, mappings: dict, allow_mf_prefix_matching: bool = False
):
    """
    Find a title match in mappings, optionally handling "MF - " prefix.

    Args:
        segment_text (str): The text segment to match
        mappings (dict): Dictionary mapping titles to keys
        allow_mf_prefix_matching (bool): If True, allows matching text without "MF - " prefix
                                       to mapping titles that have "MF - " prefix

    Returns:
        tuple: (matching_title, key) if found, (None, None) if not found
    """
    sorted_titles = sorted(mappings.keys(), key=len, reverse=True)
    segment_text_clean = segment_text.strip()

    # Check for exact match first
    for title in sorted_titles:
        if segment_text_clean.lower() == title.lower():
            return title, mappings[title]

    # If no exact match and MF prefix matching is allowed, try matching with "MF - " prefix
    if allow_mf_prefix_matching:
        for title in sorted_titles:
            # Check if the mapping title starts with "MF - " and the rest matches our segment
            if title.lower().startswith("mf - "):
                title_without_prefix = title[5:]  # Remove "MF - " prefix
                if segment_text_clean.lower() == title_without_prefix.lower():
                    return title, mappings[title]

    return None, None


def find_color_consistent_segments(paragraph):
    """
    Find consecutive runs in a paragraph that have the same text color and highlight.
    Returns a list of tuples: (start_pos, end_pos, text, colors)
    where colors is a tuple (text_color, highlight_color).
    """
    segments = []
    current_segment_start = 0
    current_segment_text = ""
    current_colors = None
    current_pos = 0

    for run in paragraph.runs:
        run_colors = get_run_colors(run)
        run_text = run.text

        if current_colors is None:
            # First run
            current_colors = run_colors
            current_segment_text = run_text
            current_segment_start = current_pos
        elif run_colors == current_colors:
            # Same colors, extend current segment
            current_segment_text += run_text
        else:
            # Different colors, save current segment and start new one
            if current_segment_text.strip():
                segments.append(
                    (
                        current_segment_start,
                        current_pos,
                        current_segment_text,
                        current_colors,
                    )
                )

            current_segment_start = current_pos
            current_segment_text = run_text
            current_colors = run_colors

        current_pos += len(run_text)

    # Add the last segment
    if current_segment_text.strip():
        segments.append(
            (current_segment_start, current_pos, current_segment_text, current_colors)
        )

    return segments


def find_non_overlapping_titles(text: str, mappings: dict):
    """
    Find the largest non-overlapping titles in a text segment.
    Returns a list of match dictionaries sorted by start position (descending).
    """
    import re

    # Find all possible title matches with their positions
    all_matches = []
    sorted_titles = sorted(mappings.keys(), key=len, reverse=True)

    for title in sorted_titles:
        pattern = re.compile(re.escape(title), flags=re.IGNORECASE)
        for match in pattern.finditer(text):
            all_matches.append(
                {
                    "start": match.start(),
                    "end": match.end(),
                    "title": title,
                    "key": mappings[title],
                    "length": len(title),
                }
            )

    # Sort by length (descending) then by start position
    all_matches.sort(key=lambda x: (-x["length"], x["start"]))

    # Select non-overlapping matches (greedy approach - largest first)
    selected_matches = []
    for match in all_matches:
        # Check if this match overlaps with any already selected match
        overlaps = False
        for selected in selected_matches:
            if not (
                match["end"] <= selected["start"] or match["start"] >= selected["end"]
            ):
                overlaps = True
                break

        if not overlaps:
            selected_matches.append(match)

    # Sort by start position (descending) for replacement
    selected_matches.sort(key=lambda x: x["start"], reverse=True)
    return selected_matches


def replace_titles_with_mergefields_respecting_colors(paragraph, mappings: dict) -> str:
    """
    Replaces titles with MERGEFIELD keys based on color consistency rules:
    - For highlighting: only if the entire title has consistent highlighting
    - For text coloring: find largest non-overlapping titles within color-consistent segments
    """
    import re

    # Get the full paragraph text
    full_text = paragraph.text

    # Find color-consistent segments
    segments = find_color_consistent_segments(paragraph)

    # Create a list of valid title matches
    valid_matches = []

    for segment_start, segment_end, segment_text, colors in segments:
        text_color, highlight_color = colors

        if highlight_color:
            # For highlighting: use the original logic (entire segment must match a title)
            segment_text_clean = segment_text.strip()

            # Check if highlight is 00FF00 (green), which allows MF prefix matching
            allow_mf_prefix = highlight_color.upper() == "00FF00"

            matching_title, key = find_title_match_with_mf_prefix(
                segment_text_clean, mappings, allow_mf_prefix
            )

            if matching_title:
                # Add TextInput: prefix for green highlights
                final_key = f"TextInput:{key}" if allow_mf_prefix else key

                valid_matches.append(
                    {
                        "start": segment_start,
                        "end": segment_end,
                        "title": matching_title,  # Use the full title from mappings
                        "key": final_key,
                        "length": len(
                            segment_text_clean
                        ),  # Length of actual text in document
                    }
                )

        elif text_color:
            # For text coloring: find largest non-overlapping titles within the segment
            segment_matches = find_non_overlapping_titles(segment_text, mappings)

            # Adjust positions to be relative to the full paragraph
            for match in segment_matches:
                valid_matches.append(
                    {
                        "start": segment_start + match["start"],
                        "end": segment_start + match["end"],
                        "title": match["title"],
                        "key": match["key"],
                        "length": match["length"],
                    }
                )

    # Sort matches by start position (descending) to replace from end to beginning
    valid_matches.sort(key=lambda x: x["start"], reverse=True)

    # Remove any overlapping matches that might have been created across segments
    final_matches = []
    for match in valid_matches:
        overlaps = False
        for final in final_matches:
            if not (match["end"] <= final["start"] or match["start"] >= final["end"]):
                overlaps = True
                break

        if not overlaps:
            final_matches.append(match)

    # Apply replacements
    result_text = full_text
    for match in final_matches:
        start_pos = match["start"]
        end_pos = match["end"]
        mergefield = f"{{ MERGEFIELD {match['key']} }}"

        # Replace the text at the specific position
        result_text = result_text[:start_pos] + mergefield + result_text[end_pos:]

    return result_text


def replace_titles_with_mergefields(text: str, mappings: dict) -> str:
    """
    Replaces all occurrences of titles in the text with their corresponding key, formatted as { MERGEFIELD key }.
    Note: This function doesn't respect color consistency. Use replace_titles_with_mergefields_respecting_colors
    for paragraph objects when color consistency is required.
    """
    import re

    # Sort titles by length descending to avoid partial replacements
    sorted_titles = sorted(mappings.keys(), key=len, reverse=True)
    for title in sorted_titles:
        # Use case-insensitive word boundaries
        pattern = re.compile(re.escape(title), flags=re.IGNORECASE)
        key = mappings[title]
        mergefield = f"{{ MERGEFIELD {key} }}"
        text = pattern.sub(mergefield, text)
    return text


def transform_text_with_if_betingelse(text: str, mappings: dict) -> str:
    """
    Applies IF Betingelse transformation to the text.
    """
    import re

    pattern = re.compile(
        r'(?i)\s+if\s+betingelse\s+(.+?)\s*(?=[“”"])[“”"]([^“”"]*)[“”"]\s*else\s*[“”"]([^“”"]*)[“”"]'
    )

    def replace_match(match):
        key, quote1, quote2 = match.groups()
        return f'{{ IF "J" = {key}" " {quote1}" " {quote2}" }}'

    return pattern.sub(replace_match, text)


def transform_text_with_single_if_condition(text: str, mappings: dict) -> str:
    """
    Transforms sentences into an IF condition with a single insert for the true case.
    """
    import re

    pattern = re.compile(r'(?i)Else til if betingelse\s+(.+?)\s*[“”"]([^“”"]*)[“”"]')

    def replace_match(match):
        key, quote = match.groups()
        return f'{{ IF "J" = {key} " {quote}" }}'

    return pattern.sub(replace_match, text)
