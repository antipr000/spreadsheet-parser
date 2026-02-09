"""
Phase 1 — Planning Agent.

Sends a multimodal request (screenshot + structural summary) to the
LLM and parses the response into a list of PlannedBlock objects.
"""

from __future__ import annotations

import logging
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from ai.factory import get_decision_service, get_decision_for_media_service
from ai.response_parser import parse_llm_json
from dto.cell_data import CellData
from dto.coordinate import BoundingBox

from agentic_flow.cell_reader import (
    read_all_cells,
    build_grid,
    build_merge_map,
)
from agentic_flow.dto.plan import PlannedBlock, TableHints
from agentic_flow.prompts.planner import get_planner_prompt
from agentic_flow.screenshot import render_sheet_screenshot
from agentic_flow.summarizer import summarise_sheet

logger = logging.getLogger(__name__)


class PlannerAgent:
    """
    Analyses a worksheet and produces a block plan.
    """

    def plan(
        self,
        ws: Worksheet,
        wb: Workbook,
        xlsx_path: str,
        computed_values: Optional[Dict[Tuple[str, str], Any]] = None,
    ) -> List[PlannedBlock]:
        """
        Run the planner on a single worksheet.

        Returns an ordered list of PlannedBlock objects.
        """
        sheet_name = ws.title or "Sheet"
        logger.info("  [Planner] Analysing sheet: %s", sheet_name)

        # 1. Read cells & build grid
        all_cells, min_row, min_col, max_row, max_col = read_all_cells(
            ws, computed_values
        )
        if not all_cells:
            logger.info("  [Planner] Sheet is empty")
            return []

        grid = build_grid(all_cells)

        # 2. Build structural summary
        summary = summarise_sheet(grid, ws, min_row, min_col, max_row, max_col)
        logger.info(
            "  [Planner] Summary: %d chars (~%d tokens)",
            len(summary),
            len(summary) // 4,
        )

        # 3. Build prompt
        prompt = get_planner_prompt(summary)

        # 4. Render screenshot
        screenshot = render_sheet_screenshot(xlsx_path, sheet_name)

        # 5. Call LLM (multimodal if screenshot available, text-only otherwise)
        raw_response: str
        if screenshot is not None:
            logger.info("  [Planner] Sending multimodal request (screenshot + text)")
            ai = get_decision_for_media_service()
            raw_response = ai.get_decision_for_media(
                prompt, screenshot, mime_type="image/png"
            )
        else:
            logger.info("  [Planner] Sending text-only request (no screenshot)")
            ai = get_decision_service()
            raw_response = ai.get_decision(prompt)

        # 6. Parse response
        parsed = parse_llm_json(raw_response)
        if parsed is None:
            logger.warning("  [Planner] Failed to parse LLM response")
            return []

        blocks_raw: List[dict]
        if isinstance(parsed, dict):
            blocks_raw = parsed.get("blocks", [])
        elif isinstance(parsed, list):
            blocks_raw = parsed
        else:
            logger.warning("  [Planner] Unexpected response type: %s", type(parsed))
            return []

        # 7. Convert to PlannedBlock DTOs
        planned: List[PlannedBlock] = []
        for i, item in enumerate(blocks_raw):
            try:
                pb = self._parse_block(item, fallback_id=f"b{i}")
                planned.append(pb)
            except Exception as exc:
                logger.warning(
                    "  [Planner] Skipping invalid block %d: %s", i, exc
                )

        logger.info("  [Planner] Identified %d block(s)", len(planned))
        return planned

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _parse_block(raw: dict, fallback_id: str) -> PlannedBlock:
        """Convert a raw dict from the LLM into a PlannedBlock."""
        block_id = raw.get("block_id", fallback_id)
        block_type = raw.get("block_type", "table")
        description = raw.get("description", "")

        bbox_raw = raw.get("bounding_box", {})
        bbox = BoundingBox(
            top_left=bbox_raw.get("top_left", "A1"),
            bottom_right=bbox_raw.get("bottom_right", "A1"),
        )

        table_hints: Optional[TableHints] = None
        th_raw = raw.get("table_hints")
        if th_raw and isinstance(th_raw, dict):
            table_hints = TableHints(
                has_multi_level_headers=th_raw.get("has_multi_level_headers", False),
                header_row_count=th_raw.get("header_row_count", 1),
                has_row_groups=th_raw.get("has_row_groups", False),
                row_group_label_column=th_raw.get("row_group_label_column"),
                merged_group_columns=th_raw.get("merged_group_columns", []),
            )

        return PlannedBlock(
            block_id=block_id,
            block_type=block_type,
            bounding_box=bbox,
            description=description,
            table_hints=table_hints,
        )
