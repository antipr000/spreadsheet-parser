"""
Top-level output DTOs for the final JSON document model.

    WorkbookResult
      └─ sheets: List[SheetResult]
           └─ chunks: Dict[str, List[Block]]
                e.g. "block0": [HeadingBlock, TableBlock]
                     "block1": [KeyValueBlock]
"""

from __future__ import annotations

from typing import Dict, List

from pydantic import BaseModel

from dto.blocks import Block


class SheetResult(BaseModel):
    """Structured output for a single worksheet."""

    sheet_name: str
    chunks: Dict[str, List[Block]] = {}


class WorkbookResult(BaseModel):
    """Top-level output for an entire workbook."""

    file_name: str
    sheets: List[SheetResult] = []
