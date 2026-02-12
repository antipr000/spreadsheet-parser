"""
LLM prompts used by the AI-assisted detection path for each block type.

Cell data is automatically sampled for large regions to stay within
token limits.
"""

from __future__ import annotations

from typing import List

from dto.cell_data import CellData
from prompts.bounding_box import _sample_cells_for_prompt


def _cells_block(cells: List[CellData]) -> str:
    """Compact multi-line representation of a cell list for LLM prompts.

    Automatically samples large cell lists to stay within token limits.
    """
    return _sample_cells_for_prompt(cells)


# -------------------------------------------------------------------
# Heading
# -------------------------------------------------------------------

def get_heading_detection_prompt(cells: List[CellData]) -> str:
    return f"""You are a spreadsheet analyst. You are given cell data for a small region of an Excel sheet.
Determine whether this region is a **heading / section title**.

A heading region typically:
- Spans 1–3 rows
- Has bold text, larger font size, or is merged across columns
- Contains short label-like text (not data values)
- Acts as a title for content below it

Cell data:
{_cells_block(cells)}

If this IS a heading, respond with a JSON object:
{{"is_heading": true, "text": "<the heading text>"}}

If this is NOT a heading, respond with:
{{"is_heading": false}}

Output ONLY the JSON object, no other text.
"""


# -------------------------------------------------------------------
# Key-Value
# -------------------------------------------------------------------

def get_key_value_detection_prompt(cells: List[CellData]) -> str:
    return f"""You are a spreadsheet analyst. You are given cell data for a region of an Excel sheet.
Determine whether this region is a **key-value / form-like** layout.

A key-value region typically:
- Has label cells on the left and corresponding value cells on the right
- Labels are short text (field names), values are data (numbers, dates, names)
- Rows are independent pairs, not part of a columnar table
- May have 2–4 columns (key, value, sometimes with gaps or units)

Cell data:
{_cells_block(cells)}

If this IS a key-value region, respond with a JSON object:
{{
  "is_key_value": true,
  "pairs": [
    {{"key_coordinate": "A1", "value_coordinate": "B1"}},
    {{"key_coordinate": "A2", "value_coordinate": "B2"}}
  ]
}}
Where each pair maps a key cell coordinate to its value cell coordinate.

If this is NOT a key-value region, respond with:
{{"is_key_value": false}}

Output ONLY the JSON object, no other text.
"""


# -------------------------------------------------------------------
# Text / Notes
# -------------------------------------------------------------------

def get_text_detection_prompt(cells: List[CellData]) -> str:
    return f"""You are a spreadsheet analyst. You are given cell data for a region of an Excel sheet.
Determine whether this region is a **free-text / notes** block.

A text/notes region typically:
- Contains sentence-length or paragraph-length prose
- Is not structured as a table (no repeating columnar pattern)
- Is not a heading (more than a short label)
- May be a disclaimer, footnote, instruction, or comment

Cell data:
{_cells_block(cells)}

If this IS a text/notes block, respond with a JSON object:
{{"is_text": true, "text": "<the full text content>"}}

If this is NOT a text/notes block, respond with:
{{"is_text": false}}

Output ONLY the JSON object, no other text.
"""


# -------------------------------------------------------------------
# Table
# -------------------------------------------------------------------

def get_table_detection_prompt(cells: List[CellData]) -> str:
    return f"""You are a data analyst. You are given cell data for a region of an Excel sheet.
There may be one or more tables in this region.

Your task:
1. Identify each distinct table.
2. For each table, determine its bounding box and which rows/columns are headers, body, and footers.

Cell data (one cell per line, only non-empty properties shown):
{_cells_block(cells)}

Rules:
- A header row is typically bold, has a distinct background color, or contains column labels.
- A footer row is typically at the bottom and may contain totals, sums, or summary formulas.
- Body rows contain the actual data values.
- Header, body, and footer row sets must NOT overlap for a single table.
- The bounding box must tightly enclose all header + body + footer cells.

Output ONLY a JSON object with this structure:
{{
  "is_table": true,
  "tables": [
    {{
      "top_left": "A1",
      "bottom_right": "D10",
      "header_rows": [1, 2],
      "header_columns": ["A", "B", "C", "D"],
      "footer_rows": [10],
      "footer_columns": ["A", "B", "C", "D"],
      "body_rows": [3, 4, 5, 6, 7, 8, 9],
      "body_columns": ["A", "B", "C", "D"]
    }}
  ]
}}

If this region does NOT contain any table, respond with:
{{"is_table": false}}

Output ONLY the JSON object, no other text.
"""
