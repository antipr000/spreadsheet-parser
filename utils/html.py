"""
Utility to render a TableBlock's cell data into an HTML <table> string.
"""

from __future__ import annotations

from typing import List

from dto.cell_data import CellData
from openpyxl.utils import column_index_from_string


def _group_cells_into_rows(cells: List[CellData]) -> List[List[CellData]]:
    """Group a flat list of cells into rows, sorted by (row, col)."""
    if not cells:
        return []

    def _sort_key(c: CellData):
        col_str = "".join(ch for ch in c.coordinate if ch.isalpha())
        row_num = int("".join(ch for ch in c.coordinate if ch.isdigit()) or "0")
        col_num = column_index_from_string(col_str) if col_str else 0
        return (row_num, col_num)

    sorted_cells = sorted(cells, key=_sort_key)

    rows: List[List[CellData]] = []
    current_row_num = None
    current_row: List[CellData] = []
    for cell in sorted_cells:
        row_num = int("".join(ch for ch in cell.coordinate if ch.isdigit()) or "0")
        if row_num != current_row_num:
            if current_row:
                rows.append(current_row)
            current_row = [cell]
            current_row_num = row_num
        else:
            current_row.append(cell)
    if current_row:
        rows.append(current_row)

    return rows


def _escape_html(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def render_table_html(
    heading: List[CellData],
    data: List[CellData],
    footer: List[CellData],
) -> str:
    """
    Render heading / data / footer cell lists into an HTML ``<table>``
    string.
    """
    parts: List[str] = ['<table border="1" cellpadding="5" cellspacing="0">']

    # <thead>
    head_rows = _group_cells_into_rows(heading)
    if head_rows:
        parts.append("  <thead>")
        for row in head_rows:
            parts.append("    <tr>")
            for cell in row:
                val = _escape_html(cell.value or "")
                parts.append(f"      <th>{val}</th>")
            parts.append("    </tr>")
        parts.append("  </thead>")

    # <tbody>
    body_rows = _group_cells_into_rows(data)
    if body_rows:
        parts.append("  <tbody>")
        for row in body_rows:
            parts.append("    <tr>")
            for cell in row:
                val = _escape_html(cell.value or "")
                parts.append(f"      <td>{val}</td>")
            parts.append("    </tr>")
        parts.append("  </tbody>")

    # <tfoot>
    foot_rows = _group_cells_into_rows(footer)
    if foot_rows:
        parts.append("  <tfoot>")
        for row in foot_rows:
            parts.append("    <tr>")
            for cell in row:
                val = _escape_html(cell.value or "")
                parts.append(f"      <td>{val}</td>")
            parts.append("    </tr>")
        parts.append("  </tfoot>")

    parts.append("</table>")
    return "\n".join(parts)
