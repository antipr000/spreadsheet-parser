"""
Spreadsheet parser — CLI entry point.

Usage:
    python parser.py <excel_file> [--output <output.json>]

Loads an Excel workbook, extracts all semantic blocks from every sheet
(headings, tables, key-value regions, text, charts), groups them into
chunks, and writes the result as a single JSON file.
"""

from __future__ import annotations

import argparse
import logging
import os
import re
import sys
from pathlib import Path
from typing import Any, Dict, Tuple

import dotenv
import formulas
import numpy as np
import openpyxl

from dto.blocks import TableBlock, Block
from dto.output import SheetResult, WorkbookResult
from extractors.sheet import SheetExtractor
from grouping import group_blocks_into_chunks
from utils.html import render_table_html

dotenv.load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(name)s  %(message)s",
)
logger = logging.getLogger(__name__)


# -------------------------------------------------------------------
# Post-processing: fill in HTML for TableBlocks
# -------------------------------------------------------------------


def _enrich_blocks(blocks: list[Block]) -> list[Block]:
    """
    Apply any post-processing enrichment to blocks before serialisation.
    Currently: render HTML for table blocks.
    """
    for block in blocks:
        if isinstance(block, TableBlock) and not block.html:
            block.html = render_table_html(
                heading=block.heading,
                data=block.data,
                footer=block.footer,
            )
    return blocks


# -------------------------------------------------------------------
# Main pipeline
# -------------------------------------------------------------------


def _compute_formula_values(file_path: str) -> Dict[Tuple[str, str], Any]:
    """
    Use the ``formulas`` library to evaluate every formula in the workbook
    and return a lookup:  ``(sheet_name_upper, cell_coordinate) → value``.

    Sheet names are normalised to uppercase for case-insensitive matching.
    """
    computed: Dict[Tuple[str, str], Any] = {}
    try:
        xl_model = formulas.ExcelModel().loads(file_path).finish()
        results = xl_model.calculate()

        # results keys look like  "'[file.xlsx]SHEET NAME'!E2"  or range variants
        file_stem = Path(file_path).name
        pattern = re.compile(
            r"'\[" + re.escape(file_stem) + r"\](.+?)'!([A-Z]+\d+)$",
            re.IGNORECASE,
        )
        for key, val in results.items():
            m = pattern.match(str(key))
            if not m:
                continue
            sheet = m.group(1).upper()
            coord = m.group(2).upper()

            # Unwrap numpy scalars / 1-element arrays
            v = val
            if hasattr(v, "value"):
                # Ranges object — get the underlying array
                v = getattr(v, "value", v)
            if isinstance(v, np.ndarray):
                if v.size == 1:
                    v = v.flat[0]
                else:
                    continue  # skip multi-cell range results
            if isinstance(v, (np.integer, np.floating)):
                v = v.item()

            computed[(sheet, coord)] = v
    except Exception:
        logger.warning(
            "Formula evaluation failed — computed values will be unavailable",
            exc_info=True,
        )

    return computed


def parse_workbook(file_path: str) -> WorkbookResult:
    """
    Parse an Excel workbook and return a structured ``WorkbookResult``.
    """
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

    extractor = SheetExtractor()
    sheet_results: list[SheetResult] = []

    for sheet_name in workbook.sheetnames:
        if sheet_name != "Sheet_2":
            continue
        logger.info("Processing sheet: %s", sheet_name)
        ws = workbook[sheet_name]

        try:
            # Extract all blocks
            blocks = extractor.extract(ws, workbook, computed_values=computed_values)

            # Enrich (e.g. render HTML for tables)
            blocks = _enrich_blocks(blocks)

            # Group headings with their content blocks into chunks
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
                "Failed to process sheet '%s' — adding empty result", sheet_name
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
        description="Parse an Excel workbook into a structured JSON document.",
    )
    parser.add_argument(
        "excel_file",
        help="Path to the .xlsx file to parse",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="Output JSON file path (default: <input_name>_output.json)",
    )
    args = parser.parse_args()

    excel_path = args.excel_file
    if not os.path.isfile(excel_path):
        logger.error("File not found: %s", excel_path)
        sys.exit(1)

    # Determine output path
    if args.output:
        output_path = args.output
    else:
        stem = Path(excel_path).stem
        output_path = f"{stem}_output.json"

    # Run pipeline
    result = parse_workbook(excel_path)

    # Serialize to JSON
    json_str = result.model_dump_json(indent=2, exclude_none=True)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(json_str)

    logger.info("Output written to %s", output_path)


if __name__ == "__main__":
    main()
