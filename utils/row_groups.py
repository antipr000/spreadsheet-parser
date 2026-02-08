"""
Row-group detection for TableBlocks.

Provides a heuristic pre-check to decide whether a table likely has row
groupings, and an LLM-based detector that identifies the hierarchy.

Tables without groupings are left untouched (row_groups stays empty).
"""

from __future__ import annotations

import logging
from typing import Dict, List, Optional, Tuple

from openpyxl.utils import column_index_from_string

from ai.factory import get_decision_service
from ai.response_parser import parse_llm_json
from dto.blocks import RowGroup, TableBlock
from dto.cell_data import CellData
from prompts.row_groups import get_row_group_prompt

logger = logging.getLogger(__name__)


# -------------------------------------------------------------------
# Helpers
# -------------------------------------------------------------------


def _row_index(cells: List[CellData]) -> Dict[int, List[CellData]]:
    """Index cells by row number."""
    rows: Dict[int, List[CellData]] = {}
    for c in cells:
        rn = int("".join(ch for ch in c.coordinate if ch.isdigit()))
        rows.setdefault(rn, []).append(c)
    return rows


def _col_index(coord: str) -> int:
    """Extract 1-based column index from A1-style coordinate."""
    col_str = "".join(ch for ch in coord if ch.isalpha())
    return column_index_from_string(col_str)


def _meaningful_value(cell: CellData) -> bool:
    """Return True if the cell has a real displayable value."""
    if cell.value is None:
        return False
    v = str(cell.value).strip()
    return v != "" and v != "0" and v != "0.0"


# -------------------------------------------------------------------
# Heuristic pre-check
# -------------------------------------------------------------------


def _find_group_header_candidates(
    body_rows: Dict[int, List[CellData]],
) -> List[int]:
    """
    Identify rows that look like group headers by pattern:
      - The row has a single meaningful text value (the label) in the
        leftmost populated column.
      - That cell has special formatting: bold OR distinct background color.
      - The row has no or very few values in the remaining (data) columns.

    We determine "data columns" by finding the median number of filled
    data columns across all rows — group header rows have significantly
    fewer filled data columns than that median.
    """
    if not body_rows:
        return []

    # Find the leftmost column across all body rows (label column)
    all_cols: set = set()
    for cells in body_rows.values():
        for c in cells:
            all_cols.add(_col_index(c.coordinate))
    if not all_cols:
        return []
    label_col = min(all_cols)

    # For each row, count how many columns beyond the label column have
    # meaningful values.
    row_data_counts: Dict[int, int] = {}
    for rn, cells in body_rows.items():
        data_count = 0
        for c in cells:
            ci = _col_index(c.coordinate)
            if ci > label_col and _meaningful_value(c):
                data_count += 1
        row_data_counts[rn] = data_count

    # Find the median data-column count (typical for a "data row")
    counts = sorted(row_data_counts.values())
    if not counts:
        return []
    median_data = counts[len(counts) // 2]

    # A group header has very few data columns relative to the median.
    # Threshold: less than 20% of the median (or absolute max of 3).
    threshold = max(3, int(median_data * 0.2))

    candidates: List[int] = []
    for rn in sorted(body_rows):
        cells = body_rows[rn]

        # Must have a value in the label column
        label_cell = None
        for c in cells:
            if _col_index(c.coordinate) == label_col and _meaningful_value(c):
                label_cell = c
                break
        if label_cell is None:
            continue

        # Must have special formatting (bold or background color)
        has_formatting = bool(label_cell.font_bold) or bool(label_cell.background_color)
        if not has_formatting:
            continue

        # Must have very few data columns
        if row_data_counts[rn] <= threshold:
            candidates.append(rn)

    return candidates


def _might_have_row_groups(table: TableBlock) -> Tuple[bool, List[int]]:
    """
    Quick heuristic check: does this table look like it *might* have
    row groupings?

    Returns (True, candidate_rows) if the pattern is found at least twice.
    Returns (False, []) otherwise.
    """
    if not table.data:
        return False, []

    body_rows = _row_index(table.data)
    if len(body_rows) < 4:
        return False, []

    candidates = _find_group_header_candidates(body_rows)

    # Need at least 2 candidates to form a repeating pattern
    if len(candidates) < 2:
        return False, []

    return True, candidates


# -------------------------------------------------------------------
# LLM-based group detection
# -------------------------------------------------------------------


def _parse_group(
    group_dict: dict,
    cell_by_row: Dict[int, List[CellData]],
    all_group_rows: set,
) -> Optional[RowGroup]:
    """Recursively parse a group dict from the LLM response into a RowGroup."""
    label_row = group_dict.get("label_row")
    if label_row is None:
        return None

    label_row = int(label_row)
    all_group_rows.add(label_row)

    # Find the label cell (first cell with a meaningful value in that row)
    row_cells = cell_by_row.get(label_row, [])
    label_cell = None
    label_text = ""
    for c in row_cells:
        if c.value is not None and str(c.value).strip():
            label_cell = c
            label_text = str(c.value).strip()
            break

    if label_cell is None:
        return None

    # Parse children recursively
    children: List[RowGroup] = []
    for child_dict in group_dict.get("children", []):
        child = _parse_group(child_dict, cell_by_row, all_group_rows)
        if child:
            children.append(child)

    return RowGroup(
        label=label_text,
        label_cell=label_cell,
        children=children,
    )


def _assign_data_rows_to_groups(
    groups: List[RowGroup],
    cell_by_row: Dict[int, List[CellData]],
    all_group_rows: set,
) -> None:
    """
    Assign non-group rows as data_rows to the appropriate group.
    Each data row belongs to the nearest group header above it.
    """
    sorted_data_rows = sorted(r for r in cell_by_row if r not in all_group_rows)
    if not sorted_data_rows or not groups:
        return

    min_row = min(cell_by_row.keys())
    max_row = max(cell_by_row.keys())

    def _get_label_row(group: RowGroup) -> int:
        return int("".join(ch for ch in group.label_cell.coordinate if ch.isdigit()))

    def _assign_recursive(
        group_list: List[RowGroup],
        start_row: int,
        end_row: int,
    ) -> None:
        for i, group in enumerate(group_list):
            label_row = _get_label_row(group)

            # Boundary: up to next sibling's label_row or end_row
            if i + 1 < len(group_list):
                boundary = _get_label_row(group_list[i + 1])
            else:
                boundary = end_row + 1

            if group.children:
                first_child_row = _get_label_row(group.children[0])
                # Rows between this group's label and first child
                for r in sorted_data_rows:
                    if label_row < r < first_child_row and r < boundary:
                        group.data_rows.extend(
                            c for c in cell_by_row.get(r, []) if c.value is not None
                        )
                _assign_recursive(group.children, first_child_row, boundary - 1)
            else:
                for r in sorted_data_rows:
                    if label_row < r < boundary:
                        group.data_rows.extend(
                            c for c in cell_by_row.get(r, []) if c.value is not None
                        )

    _assign_recursive(groups, min_row, max_row)


# -------------------------------------------------------------------
# Main entry point
# -------------------------------------------------------------------


def detect_row_groups(table: TableBlock) -> None:
    """
    Detect and populate row groupings for a TableBlock.

    Uses a heuristic to find candidate group-header rows, then asks the
    LLM to confirm the hierarchy.  If detection fails or finds nothing,
    row_groups is left empty (the table stays flat).

    Modifies the TableBlock in place.
    """
    has_groups, candidates = _might_have_row_groups(table)
    if not has_groups:
        return

    logger.info(
        "  Table %s: %d group header candidates at rows %s",
        table.bounding_box.top_left,
        len(candidates),
        candidates,
    )

    cell_by_row = _row_index(table.data)

    # Build a compact summary for the LLM: only send candidate rows
    # and a sample of data rows, not the entire table body.
    prompt = get_row_group_prompt(table.data, candidates)
    ai = get_decision_service()

    try:
        raw = ai.get_decision(prompt)
    except Exception:
        logger.warning("LLM row group detection failed", exc_info=True)
        return

    parsed = parse_llm_json(raw)
    if not isinstance(parsed, dict):
        return

    if not parsed.get("has_groups", False):
        return

    raw_groups = parsed.get("groups", [])
    if not raw_groups:
        return

    # Parse the group hierarchy
    all_group_rows: set = set()
    groups: List[RowGroup] = []
    for g in raw_groups:
        group = _parse_group(g, cell_by_row, all_group_rows)
        if group:
            groups.append(group)

    if not groups:
        return

    # Assign data rows to groups
    _assign_data_rows_to_groups(groups, cell_by_row, all_group_rows)

    table.row_groups = groups
    logger.info(
        "  Detected %d top-level row group(s) in table %s",
        len(groups),
        table.bounding_box.top_left,
    )
