"""
Base class for all agentic-flow extractors.

Each extractor receives a PlannedBlock (type + bounding box + hints) together
with the cell grid and returns one or more Block DTOs.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from dto.blocks import Block
from dto.cell_data import CellData
from agentic_flow.dto.plan import PlannedBlock


class BaseExtractor(ABC):
    """
    Interface that every block-type extractor must implement.
    """

    @abstractmethod
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
        """
        Extract structured block(s) from the region described by *planned*.

        Args:
            planned: The planned block from Phase 1 (type, bbox, hints).
            grid: Full-sheet ``(row, col) -> CellData`` lookup.
            merge_map: ``coordinate -> top-left-of-merge`` mapping.
            ws: The openpyxl Worksheet.
            wb: The openpyxl Workbook.
            computed_values: Pre-calculated formula values.
            screenshot_bytes: Full-sheet screenshot PNG bytes (for vision).

        Returns:
            A list of Block DTOs (usually one, but may be more if the
            extractor discovers sub-blocks).
        """
        ...
