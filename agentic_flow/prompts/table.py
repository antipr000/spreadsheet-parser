"""
Prompt for the Table Extractor — Pass 1 (Structure Detection).

This prompt is sent with a compact representation of the table's header
rows, a few sample body rows, all structural/bold rows, and merged
ranges.  The LLM identifies the structure without seeing every cell.
"""

from __future__ import annotations

from typing import List

from dto.cell_data import CellData
from prompts.bounding_box import get_cell_data_prompt


def _format_cells(cells: List[CellData]) -> str:
    """One line per cell, compact format."""
    return "\n".join(get_cell_data_prompt(c) for c in cells if c.value is not None)


def get_table_structure_prompt(
    header_cells: List[CellData],
    sample_body_cells: List[CellData],
    structural_row_cells: List[CellData],
    last_rows_cells: List[CellData],
    merged_ranges_text: str,
    total_rows: int,
    total_cols: int,
    top_left: str,
    bottom_right: str,
) -> str:
    return f"""You are a spreadsheet structure analyst.  You are given partial cell data from a table region ({top_left} to {bottom_right}, {total_rows} rows x {total_cols} cols).

Only a subset of cells is shown: header rows, a few sample body rows, all structural/bold rows, and the last rows.  Your task is to determine the **structure** of this table.

## HEADER ROWS (shown in full)
{_format_cells(header_cells)}

## SAMPLE BODY ROWS
{_format_cells(sample_body_cells)}

## STRUCTURAL / BOLD ROWS (potential group labels or subtotals)
{_format_cells(structural_row_cells)}

## LAST ROWS (potential footer/totals)
{_format_cells(last_rows_cells)}

## MERGED CELL RANGES within this table
{merged_ranges_text if merged_ranges_text else "(none)"}

---

Determine:

1. **header_rows**: Which row numbers are header rows?  (The top rows that contain column labels, possibly spanning multiple levels.)
2. **header_structure**: `"single"` or `"multi_level"` — does the header have merged cells forming a parent→child hierarchy?
3. **column_groups**: If multi-level, list each parent header (merged cell range + value) and its child column letters.
4. **footer_rows**: Which row numbers (if any) at the bottom contain totals / summaries?
5. **row_group_label_column**: If there are bold single-value rows in the body that act as group labels, which column contains the label? (e.g. "B").  Null if no row groups.
6. **row_groups**: List each group label row and the range of data rows it covers.  Format: [{{"label_row": 9, "label": "Group Name", "start_row": 10, "end_row": 19}}].
7. **merged_group_columns**: Columns where merged cells span multiple body rows to indicate grouping (e.g. ["CA", "CB"]).  Empty list if none.
8. **merged_groups**: For each merged group column, list the ranges: [{{"column": "CA", "start_row": 22, "end_row": 30, "label": "Text_23"}}].

Output ONLY a JSON object with these keys.  Example:

```json
{{
  "header_rows": [2, 3],
  "header_structure": "multi_level",
  "column_groups": [
    {{"parent_range": "G2:L2", "parent_label": "Revenue", "children": ["G", "H", "I", "J", "K", "L"]}}
  ],
  "footer_rows": [419],
  "row_group_label_column": "B",
  "row_groups": [
    {{"label_row": 9, "label": "Section A", "start_row": 10, "end_row": 19}},
    {{"label_row": 21, "label": "Section B", "start_row": 22, "end_row": 47}}
  ],
  "merged_group_columns": ["CA", "CB"],
  "merged_groups": [
    {{"column": "CA", "start_row": 22, "end_row": 30, "label": "Text_23"}},
    {{"column": "CB", "start_row": 50, "end_row": 55, "label": "Text_40"}}
  ]
}}
```

If the table has no row groups, set `row_group_label_column` to null and `row_groups` to [].
If the table has no multi-level headers, set `header_structure` to "single" and `column_groups` to [].
If there are no footer rows, set `footer_rows` to [].

Output ONLY the JSON object, no other text.
"""
