from pydantic import BaseModel
from typing import List, Optional


class CellData(BaseModel):
    coordinate: str
    value: Optional[str] = None
    formula: Optional[str] = None
    background_color: Optional[str] = None
    font_color: Optional[str] = None
    font_size: Optional[int] = None
    font_name: Optional[str] = None
    font_bold: Optional[bool] = None
    font_italic: Optional[bool] = None
    font_underline: Optional[bool] = None
    font_strikethrough: Optional[bool] = None
    font_subscript: Optional[bool] = None
    font_superscript: Optional[bool] = None
    merged_with: Optional[str] = None  # top-left cell of the merge range, if merged
    data_validation: Optional[List[str]] = None  # allowed values / choices
