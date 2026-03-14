import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell

def covert_merge_index_to_range(ranges):
    out = []
    for range in ranges:
        out = out + [(xl_rowcol_to_cell(range[0], range[1]), xl_rowcol_to_cell(range[0], range[2]), range[3])]
    return(out)



def get_unrepeated_header_row(header_row, startrow, startcol):
    merge_ranges = []
    header_final = []
    i = 0
    mergestart = None
    mergevalue = ''
    for entry, entry_prev in zip(header_row, ['']+header_row[:-1] ):
        # print(f'{i=}')
        if entry == entry_prev:
            header_final.append('')
            # print(f'Continuing entry')
        else:
            if (mergestart is not None) and (i+startcol-1>mergestart):
                merge_ranges += [(startrow, mergestart, startcol+i-1, mergevalue)]
                # print(f'Logging entry')
            mergestart = startcol + i
            mergevalue = entry
            header_final.append(entry)
            # print(f'Reset entry')
        i += 1
    if (mergestart is not None) and (i+startcol-1>mergestart):
        merge_ranges += [(startrow, mergestart, startcol+i-1, mergevalue)]

    return(header_final, merge_ranges)

def create_sheet(df, writer, sheet_name, duplicate_header=False):
    b = df.copy()

    idx_names = list(b.index.names)
    if idx_names == [None]:
        idx_nlevels = 0
    elif type(b.index) in [pd.Index, pd.RangeIndex]:
        idx_nlevels = 1
    else:
        idx_nlevels = b.index.nlevels

    col_nlevels = b.columns.nlevels
    col_values = b.columns.values

    merge_ranges_all = []
    autofilter_row_start = col_nlevels-1
    autofilter_row_end = autofilter_row_start + b.shape[0]
    autofilter_col_end = b.shape[1] + idx_nlevels - 1

    workbook = writer.book
    if col_nlevels == 1:
        if idx_nlevels == 0:
            b.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            b.reset_index().to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        # print(f'{idx_nlevels=}, {idx_names=}, {col_nlevels=}, {col_values=}')

        # header_rows = pd.DataFrame([list(x) for x in col_values]).T.values
        # header_rows = [list(x) for x in header_rows]

        # import pdb;pdb.set_trace();
        # header_rows = [x.values for x in header_rows]
        startrow = 0
        
        # for header_row in header_rows:
        for col_level_i in range(col_nlevels):
            # print(f"{header_row=}")
            # import pdb;pdb.set_trace();
            header_row = b.columns.get_level_values(col_level_i).tolist()
            header_row, merge_ranges = get_unrepeated_header_row(header_row, startrow, idx_nlevels)
            # print(f"{header_row=}, {merge_ranges=}")
            merge_ranges_all += merge_ranges

            write_df = pd.DataFrame(columns=header_row)
            write_df.to_excel(writer, sheet_name=sheet_name, startrow=startrow, index=False, startcol=idx_nlevels)
            # print(f'Writing header row at Excel row {startrow} with values: {header_row}')
            startrow += 1

        if idx_nlevels > 0: # Write the index names in the last header row if they exist
            pd.DataFrame(columns = idx_names).to_excel(writer, sheet_name=sheet_name, startrow=startrow-1, index=False)
            # print(f'Writing index names at Excel row {startrow-1} with values: {idx_names}')

        b.columns = ['.'.join(x).strip('.') for x in col_values]
        
        if idx_nlevels > 0:
            b = b.reset_index()

        if duplicate_header:
            autofilter_row_end += 1 # Because the names of multiindex columns are written again

            # print(f'Writing data at Excel row {startrow} with values: {b.shape[0]} rows, {b.shape[1]} columns with header')
            b.to_excel(writer, sheet_name=sheet_name, startrow=startrow, index=False)
        else:
            # print(f'Writing data at Excel row {startrow} with values: {b.shape[0]} rows, {b.shape[1]} columns without header')
            b.to_excel(writer, sheet_name=sheet_name, startrow=startrow, index=False, header=False)

    # ------------------------------------------------------------------------
    worksheet = writer.sheets[sheet_name]
    worksheet.freeze_panes(col_nlevels, idx_nlevels)
    
    # ------------------------------------------------------------------------
    # print(merge_ranges_all)
    merge_ranges_all = covert_merge_index_to_range(merge_ranges_all)
    # print(merge_ranges_all)
    for cell_start, cell_end, cell_content in merge_ranges_all:
        worksheet.merge_range(f"{cell_start}:{cell_end}", cell_content)

    # -----------------------------------------------------------------------
    worksheet.autofilter(autofilter_row_start, 0, autofilter_row_end, autofilter_col_end)
    # print(f'Filter applied from {xl_rowcol_to_cell(autofilter_row_start,0)} to {xl_rowcol_to_cell(autofilter_row_end, autofilter_col_end)}')
    
    # -----------------------------------------------------------------------
    # Row formats
    center = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True})
    # Column formats
    percent_format = workbook.add_format({'num_format': '0.00%'})
    decimal_format = workbook.add_format({'num_format': '0.00'})
    comma_format = workbook.add_format({'num_format': '#,##0'})    

    for rownum in range(autofilter_row_end - b.shape[0]+1):
        worksheet.set_row(rownum, None, center)
    
    # worksheet.set_column('A:A', None, percent_format)
    # worksheet.set_column('B:B', None, decimal_format)
    # worksheet.set_column('C:C', None, comma_format)

    # print('-'*40)

    return(1)
