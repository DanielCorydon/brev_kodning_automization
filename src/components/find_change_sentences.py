"""
Module for finding text patterns across paragraphs in Word documents using regular expressions.
"""

import re
from docx import Document
import logging
from typing import Dict, List, Any, Pattern, Optional, Set
from io import BytesIO
import json  # Add this import to fix the NameError

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


class DocumentRegexFinder:
    """
    Class for finding and extracting text matching regex patterns across paragraphs in Word documents.
    """

    def __init__(self):
        """Initialize the DocumentRegexFinder."""
        pass

    def extract_text_from_document(self, doc_input: Any) -> str:
        """
        Extract all text from a Word document, preserving paragraph structure.

        Args:
            doc_input: Path to the Word document or a Document object

        Returns:
            A string containing all document text with paragraphs separated by newlines
        """
        try:
            if isinstance(doc_input, str):
                doc = Document(doc_input)
            elif isinstance(doc_input, BytesIO):
                doc = Document(doc_input)
            else:
                raise ValueError(
                    "Invalid input type. Expected file path or BytesIO object."
                )

            # Extract text from paragraphs, preserving paragraph breaks
            paragraphs_text = [para.text for para in doc.paragraphs]

            # Extract text from tables
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        paragraphs_text.extend([para.text for para in cell.paragraphs])

            # Join all paragraphs with newlines to maintain structure
            full_text = "\n".join(paragraphs_text)

            return full_text

        except Exception as e:
            logger.error(f"Error extracting text from document: {e}")
            raise

    def find_patterns(
        self, doc_path: str, patterns: List[Pattern]
    ) -> Dict[str, List[Dict[str, Any]]]:
        """
        Find all occurrences of the provided regex patterns in the document.

        Args:
            doc_path: Path to the Word document
            patterns: List of compiled regex patterns to search for

        Returns:
            A dictionary with pattern string as key and list of matches as values
        """
        # Extract all text from the document
        doc_text = self.extract_text_from_document(doc_path)

        results = {}
        seen_full_texts: Set[str] = set()  # To track duplicates

        # Process each regex pattern
        for pattern in patterns:
            pattern_str = pattern.pattern
            results[pattern_str] = []

            # Find all matches of the pattern in the text
            matches = pattern.finditer(doc_text)

            for match in matches:
                full_text = match.group(0)

                # Skip if this exact text has already been matched
                if full_text in seen_full_texts:
                    continue

                seen_full_texts.add(full_text)

                # Extract the groups (if any)
                groups = []
                for i in range(1, len(match.groups()) + 1):
                    if match.group(i) is not None:
                        groups.append(match.group(i))

                # Add the match data to results
                results[pattern_str].append({"fullText": full_text, "groups": groups})

        return results


def process_document_with_regexes(
    doc_path: str, regex_list: List[str]
) -> Dict[str, List[Dict[str, Any]]]:
    """
    Process a document with a list of regex patterns.

    Args:
        doc_path: Path to the Word document
        regex_list: List of regex pattern strings

    Returns:
        Dictionary with results for each pattern
    """
    # Compile all regex patterns with re.DOTALL and re.UNICODE to handle Unicode characters
    compiled_patterns = []
    for regex_str in regex_list:
        try:
            # re.DOTALL makes '.' match any character including newlines
            # re.UNICODE ensures proper handling of Unicode characters
            pattern = re.compile(regex_str, re.DOTALL | re.UNICODE)
            compiled_patterns.append(pattern)
        except re.error as e:
            logger.error(f"Invalid regex pattern '{regex_str}': {e}")
            # Continue with valid patterns

    finder = DocumentRegexFinder()
    results = finder.find_patterns(doc_path, compiled_patterns)
    # Print statement for demonstration
    print("\n--- Regex Extraction Results ---\n")
    print(f"Document: {doc_path}\n")
    print(f"Regexes: {regex_list}\n")
    print("Results:\n")
    # Decode Unicode escape sequences for proper display of Danish letters
    print(json.dumps(results, indent=2, ensure_ascii=False))
    return results


if __name__ == "__main__":
    # Example usage
    import sys

    if len(sys.argv) < 3:
        print(
            "Usage: python find_change_sentences.py <docx_file> <regex1> [regex2] ..."
        )
        sys.exit(1)

    doc_path = sys.argv[1]
    regexes = sys.argv[2:]

    results = process_document_with_regexes(doc_path, regexes)

    # Print results in a readable format
    print("\n--- Processed Results ---\n")
    print(json.dumps(results, indent=2, ensure_ascii=False))
