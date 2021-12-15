import openpyxl
import pandas as pd
from my_openpyxl import find_column, write_column


def main(bom_single_path: str, bom_hier_path: str):
    # read df from excel
    bom_single = pd.read_excel(bom_single_path, header=1)
    bom_hier = pd.read_excel(bom_hier_path, header=1)
    # columns to update
    update_columns = ['In stock', 'Besteld', 'Status', 'Besteldatum', 'Leveringsdatum', 'Palletnr.']
    # columns to check
    columns = ['Component', 'Component description', 'PartNumber', 'Size', 'Length', 'Diameter', 'Thickness']

    bom_hier_update = bom_hier.loc[:, update_columns].copy()

    for i in range(len(bom_hier)):
        # value is the same or both are na
        mask1 = (bom_single.loc[:, columns] == bom_hier.loc[i, columns]) | \
                (pd.isna(bom_single.loc[:, columns]) & pd.isna(bom_hier.loc[i, columns]))
        # reduce to single column, with True if all values in the row were True
        mask2 = mask1.all(axis=1)
        # the row of the single component bom list that's the same as the row of the hierarchical bom list
        same_row = bom_single[mask2]
        # sets the status
        try:
            bom_hier_update.loc[i, :] = same_row.loc[same_row.index[0], :]
        except IndexError:  # if the row wasn't found, write '-' in all columns
            bom_hier_update.loc[i, :] = '-'

    bom_hier.loc[:, update_columns] = bom_hier_update

    # open the excel file with openpyxl
    wb = openpyxl.load_workbook(bom_hier_path)
    ws = wb.active
    # write columns
    for col in update_columns:
        col_excel = find_column(ws, col)
        write_column(ws, col_excel, bom_hier.loc[:, col])

    wb.save(bom_hier_path)
