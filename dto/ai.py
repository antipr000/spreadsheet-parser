from pydantic import BaseModel
from typing import List


class TableSchemaDTO(BaseModel):
    top_left: str
    bottom_right: str
    header_rows: List[int]
    header_columns: List[str]
    footer_rows: List[int]
    footer_columns: List[str]
    body_rows: List[int]
    body_columns: List[str]
