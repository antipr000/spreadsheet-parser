"""
Extracts tables from a worksheet.
While extracting tables we need to consider the following cases:

1. The bounding box given by ws.calculate_dimension() is absolute including all tables. There might be multiple tables in this range.
2. Both rows and columns might have labels
3. A column can have following types:
    - string
    - number
    - date
    - data validation / allowed values / choices
4. A column can have following properties that we care about:
    - header
    - footer
    - formula
    - color / background color
    - font
5. For determining bounding boxes for 2 tables, we will use an LLM.
6. Table can have multiple header rows. A header can in turn have sub headers. We need to structure the data accurately.
"""


from typing import List
from openpyxl.worksheet.worksheet import Worksheet

from dto.coordinate import BoundingBox


class TableExtractor:
    def _generate_bounding_boxes(self, worksheet: Worksheet) -> List[BoundingBox]:
        table_data_coordinates = worksheet.calculate_dimension()
