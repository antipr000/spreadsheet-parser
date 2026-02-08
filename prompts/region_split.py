"""
LLM prompt for refining heuristic region splits.

The heuristic splitter only uses whitespace (empty rows / columns) as
separators. This prompt asks the LLM to look at the actual cell contents
within each heuristic region and identify finer-grained sub-regions that
may be adjacent without any gap.
"""

from __future__ import annotations

from typing import List, Tuple

from dto.cell_data import CellData
from prompts.bounding_box import get_cell_data_prompt


def get_region_refinement_prompt(
    all_cells: List[CellData],
    heuristic_regions: List[Tuple[str, str]],
) -> str:
    """
    Build a prompt that gives the LLM:
      1. All cell data for the sheet
      2. The heuristic region boundaries (as bounding boxes)
    And asks it to refine into more granular regions.

    Args:
        all_cells: Every non-empty cell in the sheet.
        heuristic_regions: List of (top_left, bottom_right) strings from
            the heuristic splitter.
    """

    # Compact cell dump
    cells_block = "\n".join(
        get_cell_data_prompt(c) for c in all_cells if c.value is not None
    )

    # Format heuristic regions
    regions_block = "\n".join(
        f"  - Region {i}: {tl} to {br}"
        for i, (tl, br) in enumerate(heuristic_regions)
    )

    return f"""You are an expert spreadsheet analyst. You are given:
1. All cell data from a single Excel worksheet.
2. A preliminary region split that was done purely based on whitespace gaps (empty rows and columns).

Your task: Review the heuristic regions and determine whether any of them should be **split further** into smaller, independent sub-regions. Two adjacent blocks (tables, headings, key-value sections, text notes) may have been merged into one region because there was no empty row or column gap between them.

Look for these signals that a region should be split:
- A header-like row (bold, coloured, or label text) appearing in the MIDDLE of a region — this usually marks the start of a new table.
- A sudden change in column structure (e.g., first half uses columns A–D, second half uses columns A–F).
- A sudden change in data types across rows (e.g., rows of numbers followed by rows of text labels).
- A heading/title row followed by a different kind of content below it, which itself is followed by another heading/title row.
- Different semantic content stacked vertically without gaps (e.g., a key-value form above a data table).

IMPORTANT: If a heuristic region is already correct and contains only one logical block, keep it as-is. Do NOT split a single table into multiple parts.

Cell data (one cell per line, only non-empty properties shown):
{cells_block}

Heuristic regions (based on whitespace gaps):
{regions_block}

Output a JSON array where each element represents a final region. Each region must have:
- "top_left": top-left cell coordinate (e.g. "A1")
- "bottom_right": bottom-right cell coordinate (e.g. "D10")

If a heuristic region should be kept as-is, include it unchanged.
If a heuristic region should be split, replace it with the sub-regions.
Do NOT merge separate heuristic regions together.
The sub-regions must be non-overlapping and must cover all non-empty cells.

Example — if heuristic gave one region A1:D20 but it actually contains two stacked tables:
```json
[
  {{"top_left": "A1", "bottom_right": "D10"}},
  {{"top_left": "A11", "bottom_right": "D20"}}
]
```

Output ONLY the JSON array, no other text.
"""
