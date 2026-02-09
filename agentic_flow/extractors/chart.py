"""
Chart Extractor — wraps the existing ChartExtractor and uses a
text-only LLM call to generate a natural-language description from
the extracted chart data (series, categories, type, title).

Only extracts the chart that matches the planned block's bounding box
to avoid duplicates when multiple chart blocks are planned.
"""

from __future__ import annotations

import logging
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import column_index_from_string

from ai.factory import get_decision_service
from dto.blocks import Block, ChartBlock
from dto.cell_data import CellData
from dto.chart_data import ChartData
from extractors.chart import ChartExtractor as OriginalChartExtractor

from agentic_flow.dto.plan import PlannedBlock
from agentic_flow.extractors.base import BaseExtractor
from agentic_flow.prompts.chart import get_chart_description_prompt

logger = logging.getLogger(__name__)


def _parse_coord(coord: str) -> Tuple[int, int]:
    """Parse 'E2' → (row=2, col=5)."""
    col_str = "".join(c for c in coord if c.isalpha())
    row_num = int("".join(c for c in coord if c.isdigit()) or "0")
    col_num = column_index_from_string(col_str) if col_str else 0
    return row_num, col_num


def _bboxes_overlap(
    tl1: str, br1: str, tl2: str, br2: str,
) -> bool:
    """Check if two bounding boxes overlap."""
    r1_min, c1_min = _parse_coord(tl1)
    r1_max, c1_max = _parse_coord(br1)
    r2_min, c2_min = _parse_coord(tl2)
    r2_max, c2_max = _parse_coord(br2)

    if r1_max < r2_min or r2_max < r1_min:
        return False
    if c1_max < c2_min or c2_max < c1_min:
        return False
    return True


class AgenticChartExtractor(BaseExtractor):
    """
    Extracts chart data using the existing ChartExtractor and
    generates a text description via an LLM call using the
    structured chart data.

    Only returns the chart(s) whose bounding box overlaps with the
    planned block — prevents duplicates when the planner lists
    multiple chart blocks.
    """

    def __init__(self) -> None:
        self._original = OriginalChartExtractor()

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
        # Get all charts on the sheet
        all_chart_datas = self._original.extract(ws, wb)

        # Filter to only the chart(s) whose bbox overlaps the planned block
        planned_tl = planned.bounding_box.top_left
        planned_br = planned.bounding_box.bottom_right

        matching = [
            cd for cd in all_chart_datas
            if _bboxes_overlap(
                planned_tl, planned_br,
                cd.bounding_box.top_left, cd.bounding_box.bottom_right,
            )
        ]

        blocks: List[Block] = []
        for cd in matching:
            description = self._describe_chart(cd)
            blocks.append(
                ChartBlock(
                    bounding_box=cd.bounding_box,
                    chart_data=cd,
                    description=description,
                )
            )

        # If no matching chart found, create a stub from the planner info
        if not blocks:
            blocks.append(
                ChartBlock(
                    bounding_box=planned.bounding_box,
                    description=planned.description or "Chart (no data extracted)",
                )
            )

        return blocks

    @staticmethod
    def _describe_chart(cd: ChartData) -> Optional[str]:
        """
        Build a text summary of the chart data and ask the LLM to
        produce a natural-language description.
        """
        try:
            series_names = [s.name for s in cd.series if s.name]
            prompt = get_chart_description_prompt(
                cd.title, cd.chart_type, series_names
            )

            # Append the actual data so the LLM can describe specifics
            data_lines = []
            if cd.categories:
                data_lines.append(f"Categories: {cd.categories}")
            for s in cd.series:
                vals = s.values[:20]
                suffix = f" ... ({len(s.values)} total)" if len(s.values) > 20 else ""
                data_lines.append(f"Series '{s.name}': {vals}{suffix}")
            if cd.x_axis:
                data_lines.append(f"X-axis: {cd.x_axis}")
            if cd.y_axis:
                data_lines.append(f"Y-axis: {cd.y_axis}")

            if data_lines:
                prompt += "\n\nChart data:\n" + "\n".join(data_lines)

            ai = get_decision_service()
            return ai.get_decision(prompt)
        except Exception:
            logger.warning(
                "  [ChartExtractor] Chart description failed", exc_info=True
            )
            return None
