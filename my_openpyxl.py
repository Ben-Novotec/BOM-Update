import string
from openpyxl.worksheet.worksheet import Worksheet
from pandas import Series


def find_column(ws: Worksheet, column_name: str) -> str:
    """Finds the column letter in excel that corresponds with the column name in the header"""
    col = ''
    for c in string.ascii_uppercase:
        if ws[f'{c}2'].value == column_name:
            col = c
    if col == '':
        raise NameError(f'{column_name} column not found')
    return col


def write_column(ws: Worksheet, col_excel: str, col_data: Series) -> None:
    """Write the column col_data to the excel worksheet to the column with the letter col_excel"""
    for i, v in enumerate(col_data):
        ws[f'{col_excel}{i+3}'] = v

