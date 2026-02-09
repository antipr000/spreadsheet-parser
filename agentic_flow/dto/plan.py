"""
DTOs for the planning phase output.

The PlannerAgent returns a list of PlannedBlock objects — one per
identified region in the worksheet.  For tables the planner also
provides TableHints that guide the downstream TableExtractor.
"""

from __future__ import annotations

from typing import List, Literal, Optional

from pydantic import BaseModel

from dto.coordinate import BoundingBox


# ------------------------------------------------------------------
# Table-specific hints from the planner
# ------------------------------------------------------------------

class TableHints(BaseModel):
    """Structural hints the planner provides for table blocks."""

    has_multi_level_headers: bool = False
    header_row_count: int = 1
    has_row_groups: bool = False
    row_group_label_column: Optional[str] = None   # e.g. "B"
    merged_group_columns: List[str] = []            # e.g. ["CA", "CB"]


# ------------------------------------------------------------------
# PlannedBlock
# ------------------------------------------------------------------

BLOCK_TYPES = Literal[
    "heading", "table", "key_value", "text", "chart", "image"
]


class PlannedBlock(BaseModel):
    """A block identified by the PlannerAgent in Phase 1."""

    block_id: str
    block_type: BLOCK_TYPES
    bounding_box: BoundingBox
    description: str = ""
    table_hints: Optional[TableHints] = None
