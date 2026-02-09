"""
Helper to create a single-sheet .xlsx file in memory.

Used to send the actual worksheet to Gemini instead of a screenshot.
The model receives the real Excel file — merged cells, formatting,
charts, images, formulas — everything.
"""

from __future__ import annotations

import io
import logging
from typing import Optional

import openpyxl

logger = logging.getLogger(__name__)


def create_single_sheet_xlsx(
    xlsx_path: str,
    sheet_name: str,
) -> Optional[bytes]:
    """
    Load the workbook at *xlsx_path*, keep only *sheet_name*, and
    return the resulting .xlsx as raw bytes.

    Returns None if anything goes wrong.
    """
    try:
        wb = openpyxl.load_workbook(
            xlsx_path,
            data_only=False,
            keep_links=True,
        )

        # Delete every sheet except the target
        for sn in list(wb.sheetnames):
            if sn != sheet_name:
                del wb[sn]

        buf = io.BytesIO()
        wb.save(buf)
        wb.close()

        xlsx_bytes = buf.getvalue()
        logger.info(
            "  Single-sheet xlsx for '%s': %d bytes",
            sheet_name,
            len(xlsx_bytes),
        )
        return xlsx_bytes

    except Exception:
        logger.warning(
            "Failed to create single-sheet xlsx for '%s'",
            sheet_name,
            exc_info=True,
        )
        return None
