"""
Heading Extractor — simple, no LLM needed.

Reads cells in the bounding box and concatenates non-empty values
as the heading text.
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from dto.blocks import Block, HeadingBlock
from dto.cell_data import CellData

from agentic_flow.cell_reader import parse_coord, slice_grid
from agentic_flow.dto.plan import PlannedBlock
from agentic_flow.extractors.base import BaseExtractor


class HeadingExtractor(BaseExtractor):

    def extract(
        self,
        planned: PlannedBlock,
        grid: Dict[Tuple[int, int], CellData],
        merge_map: Dict[str, str],
        ws: Worksheet,
        wb: Workbook,
        *,
        computed_values: Optional[Dict[Tuple[str, str], Any]] = None,
    ) -> List[Block]:
        bbox = planned.bounding_box
        r_min, c_min = parse_coord(bbox.top_left)
        r_max, c_max = parse_coord(bbox.bottom_right)

        sub = slice_grid(grid, r_min, c_min, r_max, c_max)
        non_empty = [
            cd for cd in sub.values()
            if cd.value is not None and cd.merged_with is None
        ]

        if not non_empty:
            return []

        # Sort in reading order (row, col)
        non_empty.sort(key=lambda cd: parse_coord(cd.coordinate))

        # Build heading text — deduplicate (merged cells can repeat values)
        seen_vals: set = set()
        text_parts: List[str] = []
        for cd in non_empty:
            val = cd.value.strip() if cd.value else ""
            if val and val not in seen_vals:
                text_parts.append(val)
                seen_vals.add(val)

        text = " ".join(text_parts)

        cells = [cd for cd in sub.values() if cd.value is not None]
        cells.sort(key=lambda cd: parse_coord(cd.coordinate))

        return [
            HeadingBlock(
                bounding_box=bbox,
                text=text,
                cells=cells,
            )
        ]
