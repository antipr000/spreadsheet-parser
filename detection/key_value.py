"""
Detector for key-value / form-like regions.

Heuristic rules:
  - Region has exactly 2 meaningful column groups (key column, value column).
    A thin gap column (empty) between them is tolerated.
  - The left column has short, *heterogeneous* label-like text (no formulas).
  - The right column has data values.
  - At least 2 such pairs exist.
  - The first row must NOT look like a table header (bold / coloured).
  - Key column values must be semantically diverse (not homogeneous instances
    of the same category like "January, February, March").
"""

from __future__ import annotations

import logging
import re
from typing import List, Optional

from ai.factory import get_decision_service
from ai.response_parser import parse_llm_json
from detection.base import Detector
from dto.blocks import Block, KeyValueBlock, KeyValuePair
from dto.cell_data import CellData
from dto.region import RegionData
from prompts.detection import get_key_value_detection_prompt

logger = logging.getLogger(__name__)


def _is_label_like(cell: CellData) -> bool:
    """Heuristic: short text, no formula, looks like a field name."""
    if cell.value is None:
        return False
    if cell.formula:
        return False
    return len(cell.value) <= 60


def _is_value_like(cell: CellData) -> bool:
    """Heuristic: has a value (any length, may be numeric or a formula)."""
    return cell.value is not None


def _looks_numeric(val: str) -> bool:
    """Return True if the string looks like a number (int, float, currency)."""
    cleaned = val.strip().lstrip("$€£").replace(",", "").replace("%", "")
    try:
        float(cleaned)
        return True
    except ValueError:
        return False


def _has_header_row(region: RegionData) -> bool:
    """
    Check whether the first row of the region looks like a table header,
    either by formatting (bold / background colour) or by content
    (the first row contains generic column labels while subsequent rows
    contain data instances under those labels).
    """
    first_row = region.min_row
    row_cells = [
        cd
        for cd in (
            region.cell_at(first_row, c)
            for c in range(region.min_col, region.max_col + 1)
        )
        if cd is not None and cd.value is not None
    ]
    if not row_cells:
        return False

    # 1) Formatting-based: all bold or all with background colour
    all_bold = all(c.font_bold for c in row_cells)
    all_bg = all(c.background_color is not None for c in row_cells)
    if all_bold or all_bg:
        return True

    # 2) Content-based: first-row values are text-only labels, and at least
    #    one column has numeric data in the remaining rows while the first
    #    row value in that column is non-numeric.  This catches unformatted
    #    headers like  "Product | Total Sales"  above  "Product A | 88300".
    if region.num_rows < 3:
        return False

    first_row_values = [c.value for c in row_cells]
    # First row should be all non-numeric text
    if any(_looks_numeric(v) for v in first_row_values if v):
        return False

    # Check each column: if the body rows (row 2+) are predominantly
    # numeric but the first row is text, the first row is a header.
    for c in range(region.min_col, region.max_col + 1):
        header_cell = region.cell_at(first_row, c)
        if header_cell is None or header_cell.value is None:
            continue
        if _looks_numeric(header_cell.value):
            continue  # header cell is numeric — not a label for this col

        body_values = []
        for r in range(region.min_row + 1, region.max_row + 1):
            cd = region.cell_at(r, c)
            if cd and cd.value is not None:
                body_values.append(cd.value)

        if not body_values:
            continue
        numeric_ratio = sum(1 for v in body_values if _looks_numeric(v)) / len(
            body_values
        )
        if numeric_ratio >= 0.6:
            return True

    # 3) Content-based: first-row values are short single-word labels and
    #    body rows contain longer / multi-word values that look like
    #    instances (e.g. "Product" header above "Product A", "Product B").
    for c in range(region.min_col, region.max_col + 1):
        header_cell = region.cell_at(first_row, c)
        if header_cell is None or header_cell.value is None:
            continue
        header_val = header_cell.value.strip()

        body_values = []
        for r in range(region.min_row + 1, region.max_row + 1):
            cd = region.cell_at(r, c)
            if cd and cd.value is not None:
                body_values.append(cd.value.strip())

        if not body_values:
            continue

        # If body values contain the header text as a prefix/substring,
        # the header is a category label (e.g. "Product" → "Product A").
        prefix_matches = sum(
            1
            for v in body_values
            if v.lower().startswith(header_val.lower())
            and v.lower() != header_val.lower()
        )
        if prefix_matches >= len(body_values) * 0.5 and len(body_values) >= 2:
            return True

    return False


def _keys_are_homogeneous(keys: List[str]) -> bool:
    """
    Return True if the key values look like homogeneous instances of the
    same category rather than heterogeneous field labels.

    Examples of homogeneous keys (table-like, NOT key-value):
      - January, February, March, ...
      - Q1, Q2, Q3, Q4
      - John Doe, Jane Smith, Bob Johnson, ...
      - 2020, 2021, 2022, 2023

    Examples of heterogeneous keys (true key-value):
      - Borrower Name, Loan Amount, Due Date, Status
      - Month, Total, Average, Notes
    """
    if len(keys) < 3:
        # Too few to judge — allow it
        return False

    # If all keys are numeric, they are homogeneous (year list, ID list, etc.)
    if all(_looks_numeric(k) for k in keys):
        return True

    # If all keys have the same word count and similar length, they're likely
    # homogeneous instances (names, month names, etc.)
    word_counts = [len(k.split()) for k in keys]
    lengths = [len(k) for k in keys]
    avg_len = sum(lengths) / len(lengths)

    # Check if the word count is identical across all keys
    same_word_count = len(set(word_counts)) == 1

    # Check if lengths are very similar (coefficient of variation < 0.4)
    if avg_len > 0:
        std_len = (sum((l - avg_len) ** 2 for l in lengths) / len(lengths)) ** 0.5
        cv = std_len / avg_len
    else:
        cv = 0.0

    if same_word_count and cv < 0.4 and len(keys) >= 4:
        return True

    # Check for sequential / pattern-based keys (months, quarters, etc.)
    _MONTH_NAMES = {
        "january",
        "february",
        "march",
        "april",
        "may",
        "june",
        "july",
        "august",
        "september",
        "october",
        "november",
        "december",
        "jan",
        "feb",
        "mar",
        "apr",
        "jun",
        "jul",
        "aug",
        "sep",
        "oct",
        "nov",
        "dec",
    }
    lower_keys = {k.lower().strip() for k in keys}
    if lower_keys.issubset(_MONTH_NAMES) and len(lower_keys) >= 3:
        return True

    # Quarter pattern: Q1, Q2, Q3, Q4
    if all(re.match(r"^Q\d$", k.strip(), re.IGNORECASE) for k in keys):
        return True

    return False


class KeyValueDetector(Detector):

    # ------------------------------------------------------------------
    # Heuristic
    # ------------------------------------------------------------------

    def detect(self, region: RegionData) -> Optional[Block]:
        # Region must have 2–4 columns (key, value, optional gap / unit)
        if region.num_cols < 2 or region.num_cols > 4:
            return None

        # Must have at least 2 rows
        if region.num_rows < 2:
            return None

        # -----------------------------------------------------------
        # If the first row looks like a table header (by formatting
        # or by content), this is a table, not a key-value form.
        # -----------------------------------------------------------
        if _has_header_row(region):
            return None

        # Identify which columns are "populated" — i.e. have data in most rows.
        # A key-value region has exactly 2 populated columns (key + value).
        non_empty_cols: List[int] = []
        populated_cols: List[int] = []
        for c in range(region.min_col, region.max_col + 1):
            col_cells = region.col_cells(c)
            filled = sum(1 for cd in col_cells if cd.value is not None)
            if filled > 0:
                non_empty_cols.append(c)
            if filled > region.num_rows * 0.5:
                populated_cols.append(c)

        if len(non_empty_cols) < 2:
            return None

        # If more than 2 columns are substantially populated, this is tabular.
        if len(populated_cols) > 2:
            return None

        key_col = non_empty_cols[0]
        val_col = non_empty_cols[-1]

        # Middle columns between key and value must be mostly empty (spacers).
        if len(non_empty_cols) > 2:
            middle_cols = non_empty_cols[1:-1]
            for mc in middle_cols:
                col_cells = region.col_cells(mc)
                filled = sum(1 for cd in col_cells if cd.value is not None)
                if filled > region.num_rows * 0.3:
                    return None

        # Walk rows and build pairs
        pairs: List[KeyValuePair] = []
        for r in range(region.min_row, region.max_row + 1):
            k_cell = region.cell_at(r, key_col)
            v_cell = region.cell_at(r, val_col)
            if k_cell and v_cell and _is_label_like(k_cell) and _is_value_like(v_cell):
                pairs.append(KeyValuePair(key=k_cell, value=v_cell))

        # Need at least 2 valid pairs, and they should cover most rows
        if len(pairs) < 2:
            return None
        if len(pairs) < region.num_rows * 0.5:
            return None

        # -----------------------------------------------------------
        # Key-value keys must be heterogeneous field labels, not
        # homogeneous instances of the same category.
        # -----------------------------------------------------------
        key_values = [p.key.value for p in pairs if p.key.value is not None]
        if _keys_are_homogeneous(key_values):
            return None

        all_cells = region.non_empty_cells
        return KeyValueBlock(
            bounding_box=region.bounding_box,
            pairs=pairs,
            cells=all_cells,
        )

    # ------------------------------------------------------------------
    # AI-assisted
    # ------------------------------------------------------------------

    def detect_with_ai(self, region: RegionData) -> Optional[Block]:
        prompt = get_key_value_detection_prompt(region.non_empty_cells)
        ai = get_decision_service()
        raw = ai.get_decision(prompt)

        parsed = parse_llm_json(raw)
        if not isinstance(parsed, dict):
            return None

        if not parsed.get("is_key_value", False):
            return None

        raw_pairs = parsed.get("pairs", [])
        if not raw_pairs:
            return None

        # Build a coord → CellData lookup from region cells
        coord_map = {c.coordinate: c for c in region.cells}

        pairs: List[KeyValuePair] = []
        for rp in raw_pairs:
            k_coord = rp.get("key_coordinate")
            v_coord = rp.get("value_coordinate")
            k_cell = coord_map.get(k_coord)
            v_cell = coord_map.get(v_coord)
            if k_cell and v_cell:
                pairs.append(KeyValuePair(key=k_cell, value=v_cell))

        if not pairs:
            return None

        return KeyValueBlock(
            bounding_box=region.bounding_box,
            pairs=pairs,
            cells=region.non_empty_cells,
        )
