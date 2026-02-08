"""
LLM prompt for confirming row groupings and determining hierarchy.

The heuristic has already identified candidate group-header rows.
The LLM's job is to confirm which are real groups and determine
parent-child nesting.
"""

from __future__ import annotations

from typing import List

from dto.cell_data import CellData
from prompts.bounding_box import get_cell_data_prompt


def get_row_group_prompt(
    body_cells: List[CellData],
    candidate_rows: List[int],
) -> str:
    """
    Build a compact prompt that tells the LLM about the candidate
    group-header rows and asks it to determine the hierarchy.

    Instead of sending every cell, we send:
      - The candidate rows (full cell data)
      - A few surrounding context rows for each candidate
    """

    # Build row → cells index
    rows_idx: dict[int, list[CellData]] = {}
    for c in body_cells:
        rn = int("".join(ch for ch in c.coordinate if ch.isdigit()))
        rows_idx.setdefault(rn, []).append(c)

    # Build compact per-row summaries for candidate rows and nearby context
    context_rows: set[int] = set()
    for cr in candidate_rows:
        # Include candidate row and 1 row before/after for context
        for offset in range(-1, 2):
            if cr + offset in rows_idx:
                context_rows.add(cr + offset)

    # Build the cell block for context rows only
    context_lines: list[str] = []
    for rn in sorted(context_rows):
        cells = rows_idx.get(rn, [])
        meaningful = [c for c in cells if c.value is not None and str(c.value).strip()]
        if not meaningful:
            context_lines.append(f"Row {rn}: (empty)")
            continue

        is_candidate = rn in candidate_rows
        marker = " ** GROUP HEADER CANDIDATE **" if is_candidate else ""
        cell_strs = [get_cell_data_prompt(c) for c in meaningful[:8]]  # cap at 8 cells per row
        if len(meaningful) > 8:
            cell_strs.append(f"... and {len(meaningful) - 8} more cells")
        context_lines.append(f"Row {rn}:{marker}")
        for cs in cell_strs:
            context_lines.append(f"  {cs}")

    context_block = "\n".join(context_lines)

    # Candidate summary
    candidate_summary = "\n".join(
        f"  Row {rn}: first cell = {_first_meaningful(rows_idx.get(rn, []))!r}"
        for rn in candidate_rows
    )

    return f"""You are a spreadsheet analyst. A table has been extracted from an Excel sheet. A heuristic has identified the following rows as potential **row group headers** — rows that act as section labels grouping the data rows beneath them.

Candidate group header rows (identified by: bold/formatted, single label value, few/no data columns):
{candidate_summary}

Context (candidate rows marked with **, plus neighboring rows for reference):
{context_block}

Total body rows in this table: {len(rows_idx)}

Your task: Determine the grouping hierarchy among these candidates.

Rules:
- Confirm which candidate rows are actually group headers (some may be false positives — e.g. subtotal rows).
- If two consecutive candidate rows appear one after the other (no data rows between them), the second is typically a CHILD (sub-group) of the first.
- If a candidate row is followed by data rows and then another candidate row, they are SIBLINGS at the same level.
- A group's data rows are all rows between its label and the next group header at the same or higher level.
- Only include confirmed group headers — skip false positives.

If there ARE valid groups, respond with:
{{
  "has_groups": true,
  "groups": [
    {{
      "label_row": 9,
      "children": [
        {{"label_row": 21, "children": []}}
      ]
    }},
    {{
      "label_row": 48,
      "children": [
        {{"label_row": 49, "children": []}}
      ]
    }}
  ]
}}

If none of the candidates are real group headers, respond with:
{{"has_groups": false}}

Output ONLY the JSON object, no other text.
"""


def _first_meaningful(cells: List[CellData]) -> str:
    """Return the value of the first meaningful cell in a row."""
    for c in cells:
        if c.value is not None and str(c.value).strip():
            return str(c.value).strip()
    return ""
