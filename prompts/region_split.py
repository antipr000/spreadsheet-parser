"""
LLM prompt for refining a single heuristic region.

The heuristic splitter only uses whitespace (empty rows / columns) as
separators. This prompt asks the LLM to look at the actual cell contents
within ONE heuristic region and decide whether it should be split into
multiple sub-regions.

For large regions the cell data is **sampled** (header rows, a few body
rows, structural/bold rows, footer rows) so the prompt stays within
token limits.
"""

from __future__ import annotations

from typing import Dict, List, Set, Tuple

from openpyxl.utils import column_index_from_string, get_column_letter

from dto.cell_data import CellData
from prompts.bounding_box import get_cell_data_prompt

# ── Sampling configuration ──────────────────────────────────────────
_MAX_HEADER_ROWS = 5
_MAX_FOOTER_ROWS = 3
_MAX_SAMPLE_BODY_ROWS = 8
_MAX_COLS_FULL_DETAIL = 20
_TARGET_CHARS = 20_000  # ~5K tokens
_MAX_CELL_VALUE_LEN = 40


def _parse_coord(coord: str) -> Tuple[int, int]:
    """Parse 'E2' → (row=2, col=5)."""
    col_str = "".join(c for c in coord if c.isalpha())
    row_num = int("".join(c for c in coord if c.isdigit()) or "0")
    col_num = column_index_from_string(col_str) if col_str else 0
    return row_num, col_num


def _compact_cell_prompt(cd: CellData) -> str:
    """
    Ultra-compact cell representation with truncated value.
    """
    parts = [f"[{cd.coordinate}]"]
    if cd.value is not None:
        val = str(cd.value)[:_MAX_CELL_VALUE_LEN]
        parts.append(f'val="{val}"')
    if cd.font_bold:
        parts.append("bold")
    if cd.merged_with:
        parts.append(f"merged→{cd.merged_with}")
    if cd.formula:
        parts.append("formula")
    if cd.background_color:
        parts.append(f"bg={cd.background_color}")
    return " | ".join(parts)


def _sample_region_cells(region_cells: List[CellData]) -> str:
    """
    Build a sampled, compact representation of the region's cells.

    For small regions (≤ 200 cells) all cells are included (with
    truncated values).  For larger regions the strategy mirrors the
    agentic summariser: header rows, sampled body rows, bold/structural
    rows, and footer rows.
    """
    if not region_cells:
        return "(empty region)"

    # ── Small region: include everything (compact) ──────────────────
    if len(region_cells) <= 200:
        lines = [_compact_cell_prompt(c) for c in region_cells]
        text = "\n".join(lines)
        # Still enforce the target size
        if len(text) <= _TARGET_CHARS:
            return text

    # ── Large region: sample ────────────────────────────────────────
    # Build a grid keyed by (row, col)
    grid: Dict[Tuple[int, int], CellData] = {}
    rows_seen: Set[int] = set()
    cols_seen: Set[int] = set()
    for cd in region_cells:
        r, c = _parse_coord(cd.coordinate)
        grid[(r, c)] = cd
        rows_seen.add(r)
        cols_seen.add(c)

    if not rows_seen:
        return "(empty region)"

    sorted_rows = sorted(rows_seen)
    min_row, max_row = sorted_rows[0], sorted_rows[-1]
    min_col = min(cols_seen)
    max_col = max(cols_seen)
    total_cols = max_col - min_col + 1

    # Column sampling step for wide regions
    col_step = max(1, total_cols // _MAX_COLS_FULL_DETAIL)

    # Header rows (first N)
    header_rows = sorted_rows[:_MAX_HEADER_ROWS]

    # Footer rows (last N)
    footer_rows = sorted_rows[-_MAX_FOOTER_ROWS:]

    # Structural / bold rows (rows where all non-empty cells are bold)
    structural_rows: List[int] = []
    header_set = set(header_rows)
    footer_set = set(footer_rows)
    for r in sorted_rows:
        if r in header_set or r in footer_set:
            continue
        row_cells = [grid[(r, c)] for c in range(min_col, max_col + 1, col_step)
                     if (r, c) in grid and grid[(r, c)].value is not None]
        if row_cells and all(cd.font_bold for cd in row_cells):
            structural_rows.append(r)

    # Sampled body rows (evenly spaced, excluding header/footer/structural)
    excluded = header_set | footer_set | set(structural_rows)
    body_rows = [r for r in sorted_rows if r not in excluded]
    if body_rows:
        step = max(1, len(body_rows) // _MAX_SAMPLE_BODY_ROWS)
        sampled_body = body_rows[::step][:_MAX_SAMPLE_BODY_ROWS]
    else:
        sampled_body = []

    # Assemble sections
    sections: List[str] = []

    def _format_rows(rows: List[int], label: str) -> None:
        if not rows:
            return
        lines: List[str] = [f"--- {label} ---"]
        for r in rows:
            for c in range(min_col, max_col + 1, col_step):
                cd = grid.get((r, c))
                if cd and cd.value is not None:
                    lines.append(_compact_cell_prompt(cd))
        sections.append("\n".join(lines))

    _format_rows(header_rows, f"HEADER ROWS ({len(header_rows)} of {len(sorted_rows)} total rows)")
    _format_rows(structural_rows[:20], "STRUCTURAL / BOLD ROWS")
    _format_rows(sampled_body, f"SAMPLED BODY ROWS ({len(sampled_body)} of {len(body_rows)} body rows)")
    _format_rows(footer_rows, "FOOTER / LAST ROWS")

    text = "\n".join(sections)

    # Final safety: truncate if still too long
    if len(text) > _TARGET_CHARS:
        text = text[:_TARGET_CHARS] + "\n... (truncated)"

    return text


def get_region_refinement_prompt(
    region_cells: List[CellData],
    top_left: str,
    bottom_right: str,
) -> str:
    """
    Build a prompt for refining a single heuristic region.

    For small regions all non-empty cells are included (with truncated
    values).  For large regions the cells are sampled to stay within
    token limits.

    Args:
        region_cells: Non-empty cells within this region.
        top_left: A1-notation top-left of the heuristic region.
        bottom_right: A1-notation bottom-right of the heuristic region.
    """

    cells_block = _sample_region_cells(region_cells)

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
