"""
Text Extractor — simple, no LLM needed.

Concatenates cell values in reading order.
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from dto.blocks import Block, TextBlock
from dto.cell_data import CellData

from agentic_flow.cell_reader import parse_coord, slice_grid
from agentic_flow.dto.plan import PlannedBlock
from agentic_flow.extractors.base import BaseExtractor


class TextExtractor(BaseExtractor):

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

        non_empty.sort(key=lambda cd: parse_coord(cd.coordinate))

        # Group by row — join cells in the same row with spaces,
        # separate rows with newlines.
        rows: Dict[int, List[str]] = {}
        for cd in non_empty:
            r, _ = parse_coord(cd.coordinate)
            rows.setdefault(r, []).append(cd.value.strip() if cd.value else "")

        text = "\n".join(
            " ".join(parts)
            for _, parts in sorted(rows.items())
        )

        cells = [cd for cd in sub.values() if cd.value is not None]
        cells.sort(key=lambda cd: parse_coord(cd.coordinate))

        return [
            TextBlock(
                bounding_box=bbox,
                text=text,
                cells=cells,
            )
        ]
