"""
LLM prompt for refining a single heuristic region.

The heuristic splitter only uses whitespace (empty rows / columns) as
separators. This prompt asks the LLM to look at the actual cell contents
within ONE heuristic region and decide whether it should be split into
multiple sub-regions.
"""

from __future__ import annotations

from typing import List

from dto.cell_data import CellData
from prompts.bounding_box import get_cell_data_prompt


def get_region_refinement_prompt(
    region_cells: List[CellData],
    top_left: str,
    bottom_right: str,
) -> str:
    """
    Build a prompt for refining a single heuristic region.

    Only non-empty cells from this region are included, keeping the
    prompt small enough for any context window.

    Args:
        region_cells: Non-empty cells within this region.
        top_left: A1-notation top-left of the heuristic region.
        bottom_right: A1-notation bottom-right of the heuristic region.
    """

    cells_block = "\n".join(get_cell_data_prompt(c) for c in region_cells)

    return f"""You are an expert spreadsheet analyst. You are given cell data for ONE region of an Excel worksheet.

This region ({top_left} to {bottom_right}) was identified by a whitespace-based heuristic that splits the sheet at fully-empty rows and columns. However, this region may actually contain multiple independent blocks (tables, headings, key-value sections, text notes) stacked together without any empty-row gap between them.

Your task: Determine whether this region should be **split into smaller sub-regions** or kept as a single region.

Look for these signals that a split is needed:
- A header-like row (bold, coloured, or label text) appearing in the MIDDLE of the region — this usually marks the start of a new table or section.
- A sudden change in column structure (e.g., first half uses columns A–D, second half uses columns A–F).
- A sudden change in data types across rows (e.g., rows of numbers followed by rows of text labels).
- A heading/title row followed by a different kind of content, then another heading/title.
- Different semantic content stacked vertically (e.g., a key-value form above a data table).

IMPORTANT:
- If this region contains only ONE logical block (one table, one key-value form, etc.), keep it as-is — do NOT split it.
- Do NOT split a single table that has sub-group headers within it (bold rows used as row groupings inside one table are normal).
- Only split when the content clearly belongs to DIFFERENT, independent blocks.

Cell data for this region (one cell per line, only non-empty properties shown):
{cells_block}

Output a JSON object:
- If no split is needed:
  {{"split": false}}
- If this region should be split:
  {{"split": true, "regions": [
    {{"top_left": "A1", "bottom_right": "D10"}},
    {{"top_left": "A11", "bottom_right": "D20"}}
  ]}}

The sub-regions must be non-overlapping and their union must cover all non-empty cells in the original region.

Output ONLY the JSON object, no other text.
"""
