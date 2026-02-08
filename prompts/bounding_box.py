from typing import List
from dto.cell_data import CellData


def get_cell_data_prompt(cell_data: CellData) -> str:
    """
    Build a compact single-line representation of a cell.
    Only includes non-null fields to keep the prompt small.
    """
    parts = [f"[{cell_data.coordinate}]"]
    if cell_data.value is not None:
        parts.append(f'val="{cell_data.value}"')
    if cell_data.formula:
        parts.append(f"formula={cell_data.formula}")
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
    if cell_data.font_name:
        parts.append(f"font={cell_data.font_name}")
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


def get_bounding_box_prompt(cell_datas: List[CellData]) -> str:
    prompt = """You are a data analyst. You are given cell data for a region of an Excel sheet.
There may be one or more tables in this region.

Your task:
1. Identify each distinct table.
2. For each table, determine its bounding box and which rows/columns are headers, body, and footers.

Cell data (one cell per line, only non-empty properties shown):
"""

    for cell_data in cell_datas:
        line = get_cell_data_prompt(cell_data)
        prompt += line + "\n"

    prompt += """
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
  {
    "top_left": "A1",
    "bottom_right": "D10",
    "header_rows": [1, 2],
    "header_columns": ["A", "B", "C", "D"],
    "footer_rows": [10],
    "footer_columns": ["A", "B", "C", "D"],
    "body_rows": [3, 4, 5, 6, 7, 8, 9],
    "body_columns": ["A", "B", "C", "D"]
  }
]
```

Output ONLY the JSON array, no other text.
"""
    return prompt
