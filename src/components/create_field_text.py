"""
Module for converting regex matches to Word field replacement text.
"""

import re
from typing import Dict, List, Any, Optional
from abc import ABC, abstractmethod
import logging

logger = logging.getLogger(__name__)


class PatternProcessor(ABC):
    """Abstract base class for processing different regex patterns."""

    @abstractmethod
    def process_match(
        self,
        full_text: str,
        groups: List[str],
        mapping: Optional[Dict[str, str]] = None,
    ) -> str:
        """
        Process a regex match and return the Word field replacement text.

        Args:
            full_text: The full matched text
            groups: List of captured groups from the regex
            mapping: Optional dictionary mapping titles to keys

        Returns:
            Word field replacement text
        """
        pass


class IfElsePatternProcessor(PatternProcessor):
    """Processor for IF-ELSE conditional patterns."""

    def _apply_title_mapping(self, text: str, mapping: Optional[Dict[str, str]]) -> str:
        """
        Apply title-to-key mapping to the given text.

        Args:
            text: Text that may contain titles to be replaced
            mapping: Dictionary mapping titles to keys

        Returns:
            Text with titles replaced by their corresponding keys
        """
        if not mapping:
            return text

        # Sort titles by length (descending) to handle overlapping matches properly
        sorted_titles = sorted(mapping.keys(), key=len, reverse=True)

        result_text = text
        for title in sorted_titles:
            if title in result_text:
                key = mapping[title]
                result_text = result_text.replace(title, key)

        return result_text

    def process_match(
        self,
        full_text: str,
        groups: List[str],
        mapping: Optional[Dict[str, str]] = None,
    ) -> str:
        """
        Convert IF-ELSE pattern to Word field text.

        Expected groups:
        - groups[0]: MIDTERORD (condition words)
        - groups[1]: TEKST1 (if-text)
        - groups[2]: TEKST2 (else-text)

        Output format: { IF "J" { MERGEFIELD <MIDTERORD>}"  " TEKST1" " TEKST2"
        """
        if len(groups) < 3:
            logger.warning(f"Expected 3 groups for IF-ELSE pattern, got {len(groups)}")
            return full_text  # Return original if not enough groups

        midterord = groups[0].strip()
        tekst1 = groups[1].strip()
        tekst2 = groups[2].strip()

        # Apply title-to-key mapping to all parts
        midterord = self._apply_title_mapping(midterord, mapping)
        tekst1 = self._apply_title_mapping(tekst1, mapping)
        tekst2 = self._apply_title_mapping(tekst2, mapping)

        # Create Word field text according to the specified format
        replacement = (
            f'{{ IF "J" = {{ MERGEFIELD {midterord}}}"  " {tekst1}" " {tekst2}"'
        )

        return replacement


class PatternProcessorRegistry:
    """Registry for mapping regex patterns to their processors."""

    def __init__(self):
        self.processors: Dict[str, PatternProcessor] = {}
        self._register_default_processors()

    def _register_default_processors(self):
        """Register default pattern processors."""
        # Register the IF-ELSE pattern processor
        if_else_pattern = r'(?i)\s+if\s+betingelse\s+(.+?)\s*(?=[“”"])[“”"]([^“”"]*)[“”"]\s*else\s*[“”"]([^“”"]*)[“”"]'
        self.register_processor(if_else_pattern, IfElsePatternProcessor())

    def register_processor(self, pattern: str, processor: PatternProcessor):
        """Register a processor for a specific regex pattern."""
        self.processors[pattern] = processor

    def get_processor(self, pattern: str) -> Optional[PatternProcessor]:
        """Get the processor for a specific regex pattern."""
        return self.processors.get(pattern)


class FieldTextGenerator:
    """Main class for generating Word field replacement text from regex matches."""

    def __init__(self):
        self.registry = PatternProcessorRegistry()

    def process_regex_results(
        self,
        regex_results: List[Dict[str, Any]],
        mapping: Optional[Dict[str, str]] = None,
    ) -> List[Dict[str, Any]]:
        """
        Process regex results and add replacement text for each match.

        Args:
            regex_results: List of match dicts
            mapping: Optional dictionary mapping titles to keys

        Returns:
            Enhanced list with replacementText added to each match
        """
        enhanced_results = []

        for match in regex_results:
            pattern = match.get("regex")
            processor = self.registry.get_processor(pattern)
            enhanced_match = match.copy()
            if processor is None:
                logger.warning(f"No processor found for pattern: {pattern}")
                enhanced_match["replacementText"] = match["fullText"]  # Fallback
            else:
                try:
                    replacement_text = processor.process_match(
                        match["fullText"], match["groups"], mapping
                    )
                    enhanced_match["replacementText"] = replacement_text
                except Exception as e:
                    logger.error(f"Error processing match: {e}")
                    enhanced_match["replacementText"] = match["fullText"]  # Fallback
            enhanced_results.append(enhanced_match)

        return enhanced_results


def create_field_text_from_regex_results(
    regex_results: List[Dict[str, Any]],
    mapping: Optional[Dict[str, str]] = None,
) -> List[Dict[str, Any]]:
    """
    Main function to process regex results and add Word field replacement text.

    Args:
        regex_results: List of match dicts
        mapping: Optional dictionary mapping titles to keys

    Returns:
        Enhanced list with replacementText added to each match
    """
    generator = FieldTextGenerator()
    return generator.process_regex_results(regex_results, mapping)
