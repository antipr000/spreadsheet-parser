"""
RegionData: the data packet that every detector receives.

It bundles the cells in the candidate region together with the grid
bounds so detectors don't need to re-parse coordinates.
"""

from __future__ import annotations

from typing import Dict, List, Tuple

from pydantic import BaseModel

from dto.cell_data import CellData
from dto.coordinate import BoundingBox


class RegionData(BaseModel):
    """Pre-computed data for a rectangular candidate region."""

    # The cells that fall inside this region (row-major order)
    cells: List[CellData]

    # Bounding box in A1-notation (e.g. top_left="A1", bottom_right="D10")
    bounding_box: BoundingBox

    # Numeric (1-based) bounds for easy iteration
    min_row: int
    min_col: int
    max_row: int
    max_col: int

    # Fast (row, col) → CellData lookup.
    # Pydantic v2 will skip validation on arbitrary types with this config.
    grid: Dict[Tuple[int, int], CellData] = {}

    model_config = {"arbitrary_types_allowed": True}

    # ------------------------------------------------------------------
    # Convenience helpers
    # ------------------------------------------------------------------

    @property
    def num_rows(self) -> int:
        return self.max_row - self.min_row + 1

    @property
    def num_cols(self) -> int:
        return self.max_col - self.min_col + 1

    @property
    def non_empty_cells(self) -> List[CellData]:
        return [c for c in self.cells if c.value is not None]

    def cell_at(self, row: int, col: int) -> CellData | None:
        return self.grid.get((row, col))

    def row_cells(self, row: int) -> List[CellData]:
        """Return all cells in a given row within this region."""
        return [
            cd
            for col in range(self.min_col, self.max_col + 1)
            if (cd := self.grid.get((row, col))) is not None
        ]

    def col_cells(self, col: int) -> List[CellData]:
        """Return all cells in a given column within this region."""
        return [
            cd
            for row in range(self.min_row, self.max_row + 1)
            if (cd := self.grid.get((row, col))) is not None
        ]
