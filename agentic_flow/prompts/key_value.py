"""
Prompt for the Key-Value Extractor.
"""

from __future__ import annotations

from typing import List

from dto.cell_data import CellData
from prompts.bounding_box import get_cell_data_prompt


def get_key_value_extraction_prompt(cells: List[CellData]) -> str:
    cells_text = "\n".join(get_cell_data_prompt(c) for c in cells if c.value is not None)
    return f"""You are a spreadsheet analyst.  You are given cell data for a key-value / form-like region of an Excel sheet.

Each row (or pair of adjacent cells) represents a key → value association:
- The **key** cell contains a label (field name).
- The **value** cell contains the corresponding data (number, date, name, etc.).

Cell data:
{cells_text}

Identify all key-value pairs.  For each pair, give the key cell coordinate and value cell coordinate.

Output a JSON object:
{{
  "pairs": [
    {{"key_coordinate": "A1", "value_coordinate": "B1"}},
    {{"key_coordinate": "A2", "value_coordinate": "B2"}}
  ]
}}

Output ONLY the JSON object, no other text.
"""
