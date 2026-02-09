"""
Prompt for the Phase 1 Planning Agent.

The planner receives a screenshot + structural summary and returns a
JSON list of blocks with types, bounding boxes, and (for tables)
structural hints.
"""

from __future__ import annotations


def get_planner_prompt(sheet_summary: str) -> str:
    """
    Build the text portion of the multimodal prompt sent to the planner
    alongside the sheet screenshot.
    """
    return f"""You are an expert spreadsheet layout analyst.  You are given:
1. A **screenshot** of a single Excel worksheet.
2. A **structural summary** of the same worksheet (cell data, merged ranges, formatting cues).

Your task: identify every **independent block** on this sheet in reading order (top-to-bottom, left-to-right) and classify each block.

---

## Block types

| Type | Description |
|------|-------------|
| `heading` | A section title or label — typically 1-3 rows, bold / large font / merged across columns, short text. |
| `table` | Tabular data with one or more header rows and body rows.  May have multi-level headers (merged cells spanning columns), footer/total rows, and **row groups** (bold single-value rows that act as group labels for the rows below them). May also have merged cells in certain columns that span multiple body rows to indicate grouping. |
| `key_value` | A form-like layout where labels appear on the left and their corresponding values on the right. Rows are independent pairs, NOT a columnar table. |
| `text` | Free-form text — notes, disclaimers, footnotes, instructions. Usually prose, not structured data. |
| `chart` | An embedded chart or graph. |
| `image` | An embedded image or picture (not a chart). |

---

## How to identify blocks

1. **Whitespace gaps** — empty rows or columns between populated regions usually separate independent blocks.
2. **Formatting changes** — a bold row after non-bold rows, a change in background colour, or a different column structure signals a new block.
3. **Merged cells** — a wide merged cell at the top of a region is usually a heading; merged cells within a table spanning multiple rows in a single column indicate row grouping.
4. **Side-by-side regions** — if two populated areas are separated by one or more empty columns, they are separate blocks (even if they share the same row range).
5. **Charts / images** — these are explicitly called out in the summary ("Charts: N").

## Table structural hints

For every block classified as `table`, also provide **table_hints**:

- `has_multi_level_headers` (bool): true if header rows contain merged cells that span multiple columns (parent/child header hierarchy).
- `header_row_count` (int): how many rows at the top of this table are headers (1, 2, 3, etc.).
- `has_row_groups` (bool): true if the table body contains bold, single-value rows that act as group labels for the rows below them.
- `row_group_label_column` (string or null): the column letter (e.g. "B") that contains the group label in those single-value rows.
- `merged_group_columns` (list of strings): column letters where merged cells span multiple body rows to indicate grouping (e.g. ["CA", "CB"]).

---

## Output format

Return a JSON object with a single key `"blocks"` containing an array.  Each element:

```json
{{
  "block_id": "b0",
  "block_type": "heading",
  "bounding_box": {{"top_left": "A1", "bottom_right": "D1"}},
  "description": "Section title: Risk Summary",
  "table_hints": null
}}
```

For tables:

```json
{{
  "block_id": "b1",
  "block_type": "table",
  "bounding_box": {{"top_left": "A2", "bottom_right": "CW419"}},
  "description": "Large data table with grouped rows and merged group columns",
  "table_hints": {{
    "has_multi_level_headers": false,
    "header_row_count": 1,
    "has_row_groups": true,
    "row_group_label_column": "B",
    "merged_group_columns": ["CA", "CB"]
  }}
}}
```

**Rules:**
- Blocks must be non-overlapping.
- Bounding boxes must be in A1 notation (e.g. "A1", "CW419").
- List blocks in reading order (top-to-bottom, then left-to-right).
- If a heading appears directly above a table, list them as **two separate blocks** (the downstream system will group them).
- `table_hints` should be `null` for non-table blocks.
- If the sheet contains charts (indicated in the summary), include a block of type `chart` with an approximate bounding box.

---

## Structural summary of the worksheet

{sheet_summary}

---

Output ONLY the JSON object. No other text.
"""
