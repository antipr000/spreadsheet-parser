"""
Agentic Excel Parser — main pipeline and CLI entry point.

Usage:
    python -m agentic_flow.pipeline <excel_file> [-o <output.json>]

Two-phase pipeline per sheet:
  1. PlannerAgent  — identifies blocks (type + bbox + hints) via LLM
  2. Orchestrator  — dispatches each block to a specialised extractor

All LLM calls use the structural text summary — no spreadsheet files
are sent to the model.
"""

from __future__ import annotations

import argparse
import logging
import os
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Tuple

import dotenv
import numpy as np
import openpyxl

from dto.blocks import Block, TableBlock
from dto.output import SheetResult, WorkbookResult
from grouping import group_blocks_into_chunks
from utils.html import render_table_html

from agentic_flow.cell_reader import (
    read_all_cells,
    build_grid,
    build_merge_map,
)
from agentic_flow.planner import PlannerAgent
from agentic_flow.orchestrator import Orchestrator

dotenv.load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(name)s  %(message)s",
)
logger = logging.getLogger(__name__)


# -------------------------------------------------------------------
# Formula computation (reused from parser.py)
# -------------------------------------------------------------------

def _compute_formula_values(
    file_path: str,
    timeout_seconds: int = 30,
) -> Dict[Tuple[str, str], Any]:
    """
    Use the ``formulas`` library to evaluate every formula in the workbook
    and return a lookup: ``(sheet_name_upper, cell_coordinate) → value``.

    Aborts gracefully if computation exceeds *timeout_seconds*.
    """
    import threading

    computed: Dict[Tuple[str, str], Any] = {}

    def _do_compute() -> Dict[Tuple[str, str], Any]:
        import formulas

        xl_model = formulas.ExcelModel().loads(file_path).finish()
        results = xl_model.calculate()

        file_stem = Path(file_path).name
        pattern = re.compile(
            r"'\[" + re.escape(file_stem) + r"\](.+?)'!([A-Z]+\d+)$",
            re.IGNORECASE,
        )
        out: Dict[Tuple[str, str], Any] = {}
        for key, val in results.items():
            m = pattern.match(str(key))
            if not m:
                continue
            sheet = m.group(1).upper()
            coord_str = m.group(2).upper()

            v = val
            if hasattr(v, "value"):
                v = getattr(v, "value", v)
            if isinstance(v, np.ndarray):
                if v.size == 1:
                    v = v.flat[0]
                else:
                    continue
            if isinstance(v, (np.integer, np.floating)):
                v = v.item()

            out[(sheet, coord_str)] = v
        return out

    # Run with a timeout using threading
    result_holder: List = [{}]
    error_holder: List = [None]

    def _run():
        try:
            result_holder[0] = _do_compute()
        except Exception as exc:
            error_holder[0] = exc

    thread = threading.Thread(target=_run, daemon=True)
    thread.start()
    thread.join(timeout=timeout_seconds)

    if thread.is_alive():
        logger.warning(
            "Formula computation timed out after %ds — skipping",
            timeout_seconds,
        )
        return {}

    if error_holder[0] is not None:
        logger.warning(
            "Formula evaluation failed — computed values will be unavailable: %s",
            error_holder[0],
        )
        return {}

    return result_holder[0]


# -------------------------------------------------------------------
# Post-processing
# -------------------------------------------------------------------

def _enrich_blocks(blocks: List[Block]) -> List[Block]:
    """Render HTML for table blocks."""
    for block in blocks:
        if isinstance(block, TableBlock):
            if not block.html:
                block.html = render_table_html(
                    heading=block.heading,
                    data=block.data,
                    footer=block.footer,
                )
    return blocks


# -------------------------------------------------------------------
# Main pipeline
# -------------------------------------------------------------------

class AgenticPipeline:
    """
    Full agentic pipeline: for each sheet, run the PlannerAgent
    followed by the Orchestrator.
    """

    def __init__(self) -> None:
        self._planner = PlannerAgent()
        self._orchestrator = Orchestrator()

    def run(self, file_path: str) -> WorkbookResult:
        logger.info("Loading workbook: %s", file_path)

        workbook = openpyxl.load_workbook(
            file_path,
            data_only=False,
            read_only=False,
            keep_links=True,
            keep_vba=True,
            rich_text=True,
        )

        # Compute formula values up-front
        logger.info("Computing formula values...")
        computed_values = _compute_formula_values(file_path)
        logger.info("  -> %d formula value(s) computed", len(computed_values))

        sheet_results: List[SheetResult] = []

        for sheet_name in workbook.sheetnames:
            logger.info("=" * 60)
            logger.info("Processing sheet: %s", sheet_name)
            ws = workbook[sheet_name]

            try:
                # Phase 1: Plan
                plan = self._planner.plan(
                    ws, workbook, file_path,
                    computed_values=computed_values,
                )

                if not plan:
                    logger.info("  No blocks identified — empty sheet result")
                    sheet_results.append(SheetResult(sheet_name=sheet_name))
                    continue

                # Read cells and build grid for extraction
                all_cells, min_row, min_col, max_row, max_col = read_all_cells(
                    ws, computed_values
                )
                grid = build_grid(all_cells)
                merge_map = build_merge_map(ws)

                # Phase 2: Extract
                blocks = self._orchestrator.extract_all(
                    plan=plan,
                    grid=grid,
                    merge_map=merge_map,
                    ws=ws,
                    wb=workbook,
                    computed_values=computed_values,
                )

                # Post-process
                blocks = _enrich_blocks(blocks)

                # Group headings with content blocks
                chunks = group_blocks_into_chunks(blocks)

                sheet_results.append(
                    SheetResult(
                        sheet_name=sheet_name,
                        chunks=chunks,
                    )
                )
                logger.info(
                    "  -> %d block(s) in %d chunk(s)",
                    sum(len(v) for v in chunks.values()),
                    len(chunks),
                )

            except Exception:
                logger.exception(
                    "Failed to process sheet '%s' — adding empty result",
                    sheet_name,
                )
                sheet_results.append(SheetResult(sheet_name=sheet_name))

        return WorkbookResult(
            file_name=Path(file_path).name,
            sheets=sheet_results,
        )


# -------------------------------------------------------------------
# CLI
# -------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Agentic Excel parser — structure-aware workbook parsing.",
    )
    parser.add_argument(
        "excel_file",
        help="Path to the .xlsx file to parse",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="Output JSON file path (default: <input_name>_agentic_output.json)",
    )
    args = parser.parse_args()

    excel_path = args.excel_file
    if not os.path.isfile(excel_path):
        logger.error("File not found: %s", excel_path)
        sys.exit(1)

    if args.output:
        output_path = args.output
    else:
        stem = Path(excel_path).stem
        output_path = f"{stem}_agentic_output.json"

    pipeline = AgenticPipeline()
    result = pipeline.run(excel_path)

    json_str = result.model_dump_json(indent=2, exclude_none=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(json_str)

    logger.info("Output written to %s", output_path)


if __name__ == "__main__":
    main()
