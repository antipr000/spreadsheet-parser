from pydantic import BaseModel
from typing import List

from dto.cell_data import CellData
from dto.coordinate import BoundingBox


class TableData(BaseModel):
    """Structured representation of a single table extracted from a worksheet."""
    bounding_box: BoundingBox
    heading: List[CellData] = []
    data: List[CellData] = []
    footer: List[CellData] = []
