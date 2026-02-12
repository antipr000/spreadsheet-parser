from typing import List
from dto.cell_data import CellData

# Maximum characters for a cell value in prompts.
_MAX_CELL_VALUE_LEN = 50

# If the total prompt body exceeds this many characters, sample cells.
_TARGET_CHARS = 24_000  # ~6K tokens


def get_cell_data_prompt(cell_data: CellData) -> str:
    """
    Build a compact single-line representation of a cell.
    Only includes non-null fields to keep the prompt small.
    Cell values are truncated to ``_MAX_CELL_VALUE_LEN`` characters.
    """
    parts = [f"[{cell_data.coordinate}]"]
    if cell_data.value is not None:
        val = str(cell_data.value)[:_MAX_CELL_VALUE_LEN]
        parts.append(f'val="{val}"')
    if cell_data.formula:
        parts.append(f"formula={str(cell_data.formula)[:_MAX_CELL_VALUE_LEN]}")
    if cell_data.background_color:
        parts.append(f"bg={cell_data.background_color}")
    if cell_data.font_bold:
        parts.append("bold")
    if cell_data.font_italic:
        parts.append("italic")
    if cell_data.font_underline:
        parts.append("underline")
    if cell_data.font_size:
        parts.append(f"size={cell_data.font_size}")
    if cell_data.font_color:
        parts.append(f"color={cell_data.font_color}")
    if cell_data.font_strikethrough:
        parts.append("strikethrough")
    if cell_data.font_subscript:
        parts.append("sub")
    if cell_data.font_superscript:
        parts.append("sup")
    if cell_data.merged_with:
        parts.append(f"merged_with={cell_data.merged_with}")
    if cell_data.data_validation:
        parts.append(f"validation=[{','.join(cell_data.data_validation)}]")
    return " | ".join(parts)


def _sample_cells_for_prompt(cell_datas: List[CellData]) -> str:
    """
    Build the cell-data block for a prompt, sampling if the region is
    too large to fit within ``_TARGET_CHARS``.

    Strategy (mirrors the agentic summariser):
      1. Try including all cells (compact, truncated values).
      2. If too large, keep header-like rows (first 5), footer-like rows
         (last 3), all bold/structural rows, and a sample of body rows.
    """
    # Fast path: build full text and check size
    lines = [get_cell_data_prompt(cd) for cd in cell_datas if cd.value is not None]
    full_text = "\n".join(lines)
    if len(full_text) <= _TARGET_CHARS:
        return full_text

    # ── Need to sample ──────────────────────────────────────────────
    from openpyxl.utils import column_index_from_string

    # Parse coordinates and group by row
    row_map: dict = {}  # row_num → [CellData]
    for cd in cell_datas:
        if cd.value is None:
            continue
        col_str = "".join(c for c in cd.coordinate if c.isalpha())
        row_num = int("".join(c for c in cd.coordinate if c.isdigit()) or "0")
        row_map.setdefault(row_num, []).append(cd)

    sorted_rows = sorted(row_map.keys())
    if not sorted_rows:
        return "(empty)"

    _HEADER_N = 5
    _FOOTER_N = 3
    _SAMPLE_N = 8

    header_rows = set(sorted_rows[:_HEADER_N])
    footer_rows = set(sorted_rows[-_FOOTER_N:])

    # Bold / structural rows
    structural_rows: set = set()
    for r in sorted_rows:
        if r in header_rows or r in footer_rows:
            continue
        cells = row_map[r]
        if cells and all(cd.font_bold for cd in cells):
            structural_rows.add(r)

    # Sampled body rows
    body_rows = [r for r in sorted_rows
                 if r not in header_rows and r not in footer_rows
                 and r not in structural_rows]
    if body_rows:
        step = max(1, len(body_rows) // _SAMPLE_N)
        sampled_body = set(body_rows[::step][:_SAMPLE_N])
    else:
        sampled_body = set()

    keep_rows = header_rows | footer_rows | structural_rows | sampled_body

    sampled_lines: List[str] = []
    total_rows_count = len(sorted_rows)
    sampled_lines.append(
        f"(Sampled {len(keep_rows)} of {total_rows_count} rows to fit token limit)"
    )
    for r in sorted_rows:
        if r not in keep_rows:
            continue
        for cd in row_map[r]:
            sampled_lines.append(get_cell_data_prompt(cd))

    text = "\n".join(sampled_lines)
    # Final safety truncation
    if len(text) > _TARGET_CHARS:
        text = text[:_TARGET_CHARS] + "\n... (truncated)"
    return text


def get_bounding_box_prompt(cell_datas: List[CellData]) -> str:
    cells_block = _sample_cells_for_prompt(cell_datas)

    prompt = f"""You are a data analyst. You are given cell data for a region of an Excel sheet.
There may be one or more tables in this region.

Your task:
1. Identify each distinct table.
2. For each table, determine its bounding box and which rows/columns are headers, body, and footers.

Cell data (one cell per line, only non-empty properties shown):
{cells_block}

Rules:
- A header row is typically bold, has a distinct background color, or contains column labels.
- A footer row is typically at the bottom and may contain totals, sums, or summary formulas.
- Body rows contain the actual data values.
- Header, body, and footer row sets must NOT overlap for a single table.
- The bounding box must tightly enclose all header + body + footer cells.

Output ONLY a JSON array. Each element must have exactly these fields:
- "top_left": top-left cell coordinate of the table (e.g. "A1")
- "bottom_right": bottom-right cell coordinate of the table (e.g. "F20")
- "header_rows": array of row numbers (integers) that are header rows
- "header_columns": array of column letters (strings) that are header columns
- "footer_rows": array of row numbers (integers) that are footer rows
- "footer_columns": array of column letters (strings) that are footer columns
- "body_rows": array of row numbers (integers) that are body/data rows
- "body_columns": array of column letters (strings) that are body/data columns

Example:
```json
[
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
```

Output ONLY the JSON array, no other text.
"""
    return prompt
