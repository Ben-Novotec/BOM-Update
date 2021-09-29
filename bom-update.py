import pandas as pd


def main(bom_single_path: str, bom_hier_path: str):
    # read df from excel
    bom_single = pd.read_excel(bom_single_path, header=1)
    bom_hier = pd.read_excel(bom_hier_path, header=1)

    bom_hier_status_up = bom_hier.loc[:, 'Status'].copy()

    columns = ['Component', 'Component description', 'PartNumber', 'Size', 'Length', 'Diameter', 'Thickness']
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
            bom_hier_status_up.loc[i] = same_row.loc[same_row.index[0], 'Status']
        except IndexError:  # if the row wasn't found, write 'assembly' as the status
            bom_hier_status_up.loc[i] = 'assembly'

    bom_hier.loc[:, 'Status'] = bom_hier_status_up
    # TODO add openpyxl like in production_bom
    bom_hier.to_excel('test.xlsx')


if __name__ == '__main__':
    main('Bill of Materials 103098 - single components.xlsx', 'Bill of Materials 103098 - hierarchical.xlsx')

