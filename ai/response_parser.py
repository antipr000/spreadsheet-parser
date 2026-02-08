"""
Generic utilities for parsing structured data from LLM responses.
"""

import json
import logging
import re
from typing import Any, Dict, List, Optional, Union

logger = logging.getLogger(__name__)


def parse_llm_json(raw: str) -> Optional[Union[Dict[str, Any], List[Any]]]:
    """
    Extract a JSON object or array from a (possibly messy) LLM response.

    Handles common issues:
      - Markdown code fences (```json ... ```)
      - Leading/trailing prose around the JSON payload
      - Nested structures

    Returns the parsed Python dict/list, or ``None`` if no valid JSON was
    found.
    """
    # Strip markdown code fences if present
    cleaned = raw.strip()
    cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
    cleaned = re.sub(r"\s*```$", "", cleaned)
    cleaned = cleaned.strip()

    # Try to locate a JSON array first, then a JSON object
    json_str = _extract_json_substring(cleaned, "[", "]")
    if json_str is None:
        json_str = _extract_json_substring(cleaned, "{", "}")
    if json_str is None:
        logger.warning(
            "LLM response did not contain a JSON object or array: %s",
            raw[:200],
        )
        return None

    try:
        return json.loads(json_str)
    except json.JSONDecodeError as exc:
        logger.warning("Failed to parse LLM JSON: %s — %s", exc, json_str[:200])
        return None


def _extract_json_substring(
    text: str, open_char: str, close_char: str
) -> Optional[str]:
    """Find the outermost balanced ``open_char … close_char`` substring."""
    start = text.find(open_char)
    end = text.rfind(close_char)
    if start == -1 or end == -1 or end <= start:
        return None
    return text[start : end + 1]
