"""
Block DTOs representing the different semantic block types that can be
detected within a worksheet region.

Every block carries a bounding_box and a flat list of cells for traceability.
Specialised fields vary by block type.
"""

from __future__ import annotations

from typing import List, Literal, Optional, Union

from pydantic import BaseModel

from dto.cell_data import CellData
from dto.coordinate import BoundingBox
from dto.chart_data import ChartData


# -------------------------------------------------------------------
# Key-value helper
# -------------------------------------------------------------------

class KeyValuePair(BaseModel):
    """A single key → value association in a form-like region."""
    key: CellData
    value: CellData


# -------------------------------------------------------------------
# Concrete block types
# -------------------------------------------------------------------

class HeadingBlock(BaseModel):
    block_type: Literal["heading"] = "heading"
    bounding_box: BoundingBox
    text: str
    cells: List[CellData] = []


class RowGroup(BaseModel):
    """A group of rows within a table, forming a hierarchy."""
    label: str
    label_cell: CellData
    data_rows: List[CellData] = []
    children: List["RowGroup"] = []


class TableBlock(BaseModel):
    block_type: Literal["table"] = "table"
    bounding_box: BoundingBox
    title: Optional[str] = None
    heading: List[CellData] = []
    data: List[CellData] = []
    footer: List[CellData] = []
    html: str = ""
    cells: List[CellData] = []
    row_groups: List[RowGroup] = []


class KeyValueBlock(BaseModel):
    block_type: Literal["key_value"] = "key_value"
    bounding_box: BoundingBox
    pairs: List[KeyValuePair] = []
    cells: List[CellData] = []


class TextBlock(BaseModel):
    block_type: Literal["text"] = "text"
    bounding_box: BoundingBox
    text: str
    cells: List[CellData] = []


class ChartBlock(BaseModel):
    block_type: Literal["chart"] = "chart"
    bounding_box: BoundingBox
    chart_data: Optional[ChartData] = None
    description: Optional[str] = None
    cells: List[CellData] = []


class ImageBlock(BaseModel):
    block_type: Literal["image"] = "image"
    bounding_box: BoundingBox
    description: Optional[str] = None
    cells: List[CellData] = []


# -------------------------------------------------------------------
# Discriminated union  (for serialisation / Pydantic parsing)
# -------------------------------------------------------------------

Block = Union[HeadingBlock, TableBlock, KeyValueBlock, TextBlock, ChartBlock, ImageBlock]
