"""
Key-Value Extractor — uses LLM to identify key-value pairs.
"""

from __future__ import annotations

import logging
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from ai.factory import get_decision_service
from ai.response_parser import parse_llm_json
from dto.blocks import Block, KeyValueBlock, KeyValuePair
from dto.cell_data import CellData

from agentic_flow.cell_reader import parse_coord, slice_grid
from agentic_flow.dto.plan import PlannedBlock
from agentic_flow.extractors.base import BaseExtractor
from agentic_flow.prompts.key_value import get_key_value_extraction_prompt

logger = logging.getLogger(__name__)


class KeyValueExtractor(BaseExtractor):

    def extract(
        self,
        planned: PlannedBlock,
        grid: Dict[Tuple[int, int], CellData],
        merge_map: Dict[str, str],
        ws: Worksheet,
        wb: Workbook,
        *,
        computed_values: Optional[Dict[Tuple[str, str], Any]] = None,
        screenshot_bytes: Optional[bytes] = None,
    ) -> List[Block]:
        bbox = planned.bounding_box
        r_min, c_min = parse_coord(bbox.top_left)
        r_max, c_max = parse_coord(bbox.bottom_right)

        sub = slice_grid(grid, r_min, c_min, r_max, c_max)
        non_empty = [cd for cd in sub.values() if cd.value is not None]

        if not non_empty:
            return []

        non_empty.sort(key=lambda cd: parse_coord(cd.coordinate))

        # Build a coord→CellData lookup for quick access
        coord_lookup: Dict[str, CellData] = {
            cd.coordinate: cd for cd in non_empty
        }

        # Ask LLM for pair identification
        pairs = self._detect_pairs_with_llm(non_empty, coord_lookup)

        # Fallback: heuristic pairing (left cell = key, right cell = value)
        if not pairs:
            pairs = self._heuristic_pairs(sub, r_min, c_min, r_max, c_max)

        cells = list(non_empty)

        return [
            KeyValueBlock(
                bounding_box=bbox,
                pairs=pairs,
                cells=cells,
            )
        ]

    def _detect_pairs_with_llm(
        self,
        cells: List[CellData],
        coord_lookup: Dict[str, CellData],
    ) -> List[KeyValuePair]:
        prompt = get_key_value_extraction_prompt(cells)
        try:
            ai = get_decision_service()
            raw = ai.get_decision(prompt)
            parsed = parse_llm_json(raw)
            if not isinstance(parsed, dict):
                return []

            pairs_raw = parsed.get("pairs", [])
            pairs: List[KeyValuePair] = []
            for p in pairs_raw:
                key_coord = p.get("key_coordinate", "")
                val_coord = p.get("value_coordinate", "")
                key_cell = coord_lookup.get(key_coord)
                val_cell = coord_lookup.get(val_coord)
                if key_cell and val_cell:
                    pairs.append(KeyValuePair(key=key_cell, value=val_cell))
            return pairs
        except Exception:
            logger.warning(
                "  [KV Extractor] LLM pair detection failed",
                exc_info=True,
            )
            return []

    @staticmethod
    def _heuristic_pairs(
        sub: Dict[Tuple[int, int], CellData],
        r_min: int,
        c_min: int,
        r_max: int,
        c_max: int,
    ) -> List[KeyValuePair]:
        """
        Simple heuristic: for each row, the first non-empty cell is the
        key and the second is the value.
        """
        pairs: List[KeyValuePair] = []
        for r in range(r_min, r_max + 1):
            row_cells = []
            for c in range(c_min, c_max + 1):
                cd = sub.get((r, c))
                if cd and cd.value is not None:
                    row_cells.append(cd)
            if len(row_cells) >= 2:
                pairs.append(KeyValuePair(key=row_cells[0], value=row_cells[1]))
        return pairs
