"""
Phase 2 — Extraction Orchestrator.

Receives the block plan from Phase 1 and dispatches each PlannedBlock
to the appropriate extractor.  Collects and returns the resulting
Block DTOs in reading order.
"""

from __future__ import annotations

import logging
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from dto.blocks import Block
from dto.cell_data import CellData

from agentic_flow.dto.plan import PlannedBlock
from agentic_flow.extractors.base import BaseExtractor
from agentic_flow.extractors.heading import HeadingExtractor
from agentic_flow.extractors.table import TableExtractor
from agentic_flow.extractors.key_value import KeyValueExtractor
from agentic_flow.extractors.text import TextExtractor
from agentic_flow.extractors.chart import AgenticChartExtractor
from agentic_flow.extractors.image import ImageExtractor

logger = logging.getLogger(__name__)


class Orchestrator:
    """
    Dispatches each PlannedBlock to the correct extractor and
    collects the results.
    """

    def __init__(self) -> None:
        self._extractors: Dict[str, BaseExtractor] = {
            "heading": HeadingExtractor(),
            "table": TableExtractor(),
            "key_value": KeyValueExtractor(),
            "text": TextExtractor(),
            "chart": AgenticChartExtractor(),
            "image": ImageExtractor(),
        }

    def extract_all(
        self,
        plan: List[PlannedBlock],
        grid: Dict[Tuple[int, int], CellData],
        merge_map: Dict[str, str],
        ws: Worksheet,
        wb: Workbook,
        *,
        computed_values: Optional[Dict[Tuple[str, str], Any]] = None,
        screenshot_bytes: Optional[bytes] = None,
    ) -> List[Block]:
        """
        Run extraction for every planned block and return all Block
        DTOs in reading order.
        """
        blocks: List[Block] = []

        for planned in plan:
            extractor = self._extractors.get(planned.block_type)
            if extractor is None:
                logger.warning(
                    "  [Orchestrator] No extractor for type '%s' — skipping block %s",
                    planned.block_type,
                    planned.block_id,
                )
                continue

            logger.info(
                "  [Orchestrator] Extracting %s block '%s' (%s → %s)",
                planned.block_type,
                planned.block_id,
                planned.bounding_box.top_left,
                planned.bounding_box.bottom_right,
            )

            try:
                result = extractor.extract(
                    planned=planned,
                    grid=grid,
                    merge_map=merge_map,
                    ws=ws,
                    wb=wb,
                    computed_values=computed_values,
                    screenshot_bytes=screenshot_bytes,
                )
                blocks.extend(result)
                logger.info(
                    "  [Orchestrator]   -> %d block(s) extracted", len(result)
                )
            except Exception:
                logger.exception(
                    "  [Orchestrator] Extraction failed for block %s — skipping",
                    planned.block_id,
                )

        return blocks
