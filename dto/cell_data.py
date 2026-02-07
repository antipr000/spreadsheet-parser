from pydantic import BaseModel
from typing import Optional


class CellData(BaseModel):
    cell_address: str
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
    font_color: Optional[str] = None
