from pydantic import BaseModel


class BoundingBox(BaseModel):
    top_left: str
    bottom_right: str
