import pandas as pd
import numpy as np


def main(bom_old_path: str, bom_new_path: str):
    # read df from excel
    bom_old = pd.read_excel(bom_old_path, header=1)
    bom_new = pd.read_excel(bom_new_path, header=1)
    # columns to update
    update_columns = ['In stock', 'Besteld', 'Besteldatum', 'Leveringsdatum', 'Serienr.', 'Tagnr.', 'Palletnr.', 'Status']
    # columns to check
    check_columns = ['Component', 'Component description', 'PartNumber', 'Size', 'Length', 'Diameter', 'Thickness']

    list_update_rows = []
    revision_series = pd.Series([np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, 'Revisie!'],
                                index=update_columns)
    for i in range(len(bom_new)):
        # value is the same or both are na
        mask1 = (bom_old.loc[:, check_columns] == bom_new.loc[i, check_columns]) | \
                (pd.isna(bom_old.loc[:, check_columns]) & pd.isna(bom_new.loc[i, check_columns]))
        # reduce to single column, with True if all values in the row were True
        mask2 = mask1.all(axis=1)
        # the row of the old bom list that's the same as the row of the new bom list
        same_row = bom_old[mask2]

        # if the row wasn't found add 'Revisie!' as Status
        if not mask2.any():
            row_series = revision_series.append(bom_new.loc[i, :])
            list_update_rows.append(row_series)
        else:
            qty_old = same_row.loc[same_row.index[0], 'Quantity']
            qty_new = bom_new.loc[i, 'Quantity']
            row_series = same_row.loc[same_row.index[0], update_columns].append(bom_new.loc[i, :])
            # if the quantity is the same in the old and new bom list
            if qty_old == qty_new:
                list_update_rows.append(row_series)
            # if the quantity is different
            else:
                # add a row with the old quantity and status
                row_series.loc['Quantity'] = qty_old
                list_update_rows.append(row_series)
                # add a row with the rest of the quantity and 'Revisie!' as status
                row_series = revision_series.append(bom_new.loc[i, :])
                row_series.loc['Quantity'] = qty_new - qty_old
                list_update_rows.append(row_series)

    df_update = pd.DataFrame(list_update_rows, index=range(len(list_update_rows)))
    df_update.to_excel(f'{bom_old_path[:-5]} - Revisie.xlsx', index=False, startrow=1)


if __name__ == '__main__':
    main('test bom 1.xlsx', 'test bom 2.xlsx')
