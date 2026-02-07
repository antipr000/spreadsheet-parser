from pydantic import BaseModel
from typing import Any, List, Optional

from dto.coordinate import BoundingBox


class DataRange(BaseModel):
    sheet_name: str
    start: str
    end: str


class ChartSeries(BaseModel):
    """A single data series in a chart (e.g. one bar group or one pie ring)."""
    name: Optional[str] = None
    data_range: Optional[DataRange] = None
    values: List[Any] = []


class ChartData(BaseModel):
    title: Optional[str] = None
    x_axis: Optional[str] = None
    y_axis: Optional[str] = None
    bounding_box: BoundingBox
    chart_type: str
    categories: List[str] = []
    category_range: Optional[DataRange] = None
    series: List[ChartSeries] = []
