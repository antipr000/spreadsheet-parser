from pydantic import BaseModel


class Coordinate(BaseModel):
    row: str
    column: str


class BoundingBox(BaseModel):
    top_left: Coordinate
    bottom_right: Coordinate
