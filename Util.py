import numpy as np
from openpyxl.utils.cell import get_column_letter, column_index_from_string

# python: 0-indexed, excel: 1-indexed
def convert_python_row_to_excel_row(row: int) -> str:
    return str(row + 1)

def convert_excel_row_to_pythonl_row(row: str) -> int:
    return int(row) - 1

# python: 0-indexed, excel: A, B, C...
def convert_python_column_to_excel_column(column: int) -> str:
    return get_column_letter(column + 1)

def convert_excel_column_to_python_column(column: int) -> str:
    return column_index_from_string(column) - 1

# color has 2 representations, rgb and theme
# rgb is aRGB string in hex, while exact value of theme is according to palette of file
# TODO: convert theme to rgb code
def get_color_value(color):
    return color.rgb if color.type == "rgb" else color.theme

# solve merge_cells problem 
def fill_merge_cell(values):
    prev = values[0]
    for i in range(1, len(values)):
        if values[i] is None:
            values[i] = prev
        else:
            prev = values[i]
    return values

def root_of_merged_cell(sheet, coord):
    """ Find the parent of the merged cell by iterating through the range of merged cells """
    # Note: if there are many merged cells in a large spreadsheet, this may become inefficient
    for merged in sheet.merged_cells.ranges:
        if coord in merged:
            return merged