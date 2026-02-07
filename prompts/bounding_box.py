from typing import List
from dto.cell_data import CellData


def get_cell_data_prompt(cell_data: CellData) -> str:
    return f"""
       - cell_address: {cell_data.cell_address}
       - value: {cell_data.value}
       - formula: {cell_data.formula}
       - background_color: {cell_data.background_color}
       - font_color: {cell_data.font_color}
       - font_size: {cell_data.font_size}
       - font_name: {cell_data.font_name}
       - font_bold: {cell_data.font_bold}
       - font_italic: {cell_data.font_italic}
       - font_underline: {cell_data.font_underline}
       - font_strikethrough: {cell_data.font_strikethrough}
       - font_subscript: {cell_data.font_subscript}
       - font_superscript: {cell_data.font_superscript}
       - font_color: {cell_data.font_color}
       - font_size: {cell_data.font_size}
       - font_name: {cell_data.font_name}
       - font_bold: {cell_data.font_bold}
       - font_italic: {cell_data.font_italic}
       - font_underline: {cell_data.font_underline}
       - font_strikethrough: {cell_data.font_strikethrough}
       - font_subscript: {cell_data.font_subscript}
       - font_superscript: {cell_data.font_superscript}
       \n\n
    """


def get_bounding_box_prompt(cell_datas: List[CellData]) -> str:
    prompt = f"""
      You are a data analyst. You are given the cell data for an excel sheet. There can be multiple tables in the sheet.
      You need to identify all the different tables that exist in the sheet and define their bounding boxes.\n
      The cell data is given in the following format: \n
        - cell_address: Address of the cell in the sheet \n
        - value: Value of the cell \n
        - formula: Formula of the cell \n
        - background_color: Background color of the cell \n
        - font_color: Font color of the cell \n
        - font_size: Font size of the cell \n
        - font_name: Font name of the cell \n
        - font_bold: Whether the font is bold \n
        - font_italic: Whether the font is italic \n
        - font_underline: Whether the font is underlined \n
        - font_strikethrough: Whether the font is strikethrough \n
        - font_subscript: Whether the font is subscript \n
        - font_superscript: Whether the font is superscript \n\n
    """

    for cell_data in cell_datas:
        prompt += get_cell_data_prompt(cell_data)

    prompt += f"""
    You need to identify all the different tables that exist in the sheet and define their bounding boxes.
    The bounding boxes are defined by the top left and bottom right coordinates of the table.
    The output should be an array consisting of json objects each containing the bounding box coordinates.
    Each bounding box json object should have the following fields: top_left, bottom_right.
    top_left and bottom_right are objects with the following fields: row, column. \n

    Example output:
    [
      {
        "top_left": {
          "row": "A1",
          "column": "B1"
        },
        "bottom_right": {
          "row": "A10",
          "column": "B10"
        }
      }
    ]
    """
    return prompt
