"""
Excel export utilities for pandas DataFrames.

This module provides enhanced Excel export functionality for pandas DataFrames,
with special handling for multi-level column headers and indexes, automatic
cell merging, borders, and professional formatting.
"""

import datetime
import numpy as np
import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell


def _is_numeric_col(series):
    for val in series:
        if pd.isna(val):
            continue
        if isinstance(val, float) and np.isinf(val):
            continue
        if not isinstance(val, (int, float, np.integer, np.floating)):
            return False
    return True


def _get_decimal_places(series):
    numeric = pd.to_numeric(series, errors='coerce')
    finite = numeric.dropna()
    finite = finite[np.isfinite(finite)]
    if len(finite) == 0:
        return 0
    abs_vals = finite.abs()
    is_whole = ((finite % 1) == 0).all()
    p75 = abs_vals.quantile(0.75)
    if is_whole or p75 > 1000:
        return 0
    elif p75 > 1:
        return 2
    else:
        def _leading_zeros(v):
            v = abs(float(v))
            if v == 0 or v >= 0.1:
                return 0
            n = 0
            while v < 0.1:
                v *= 10
                n += 1
            return n
        return int(finite.apply(_leading_zeros).quantile(0.75)) + 2


def _is_pct_col_name(col_name):
    s = str(col_name)
    return '%' in s or any(p in s.lower() for p in ['perc', 'pct'])


def convert_merge_index_to_range(ranges):
    """
    Convert merge range indices to Excel cell range notation.

    Takes a list of merge range tuples containing row/column indices and
    converts them to Excel cell references (e.g., 'A1:C1').

    Args:
        ranges (list): List of tuples in format (row, start_col, end_col, value)
                      where row, start_col, end_col are 0-indexed integers

    Returns:
        list: List of tuples in format (start_cell, end_cell, value)
              where cells are Excel notation strings (e.g., 'A1', 'C1')

    Example:
        >>> convert_merge_index_to_range([(0, 1, 3, 'Sales')])
        [('B1', 'D1', 'Sales')]
    """
    out = []
    for merge_range in ranges:
        start_cell = xl_rowcol_to_cell(merge_range[0], merge_range[1])
        end_cell = xl_rowcol_to_cell(merge_range[0], merge_range[2])
        value = merge_range[3]
        out.append((start_cell, end_cell, value))
    return out


def get_unrepeated_header_row(header_row, startrow, startcol):
    """
    Process a header row to identify consecutive duplicate values for merging.

    This function analyzes a header row and identifies sequences of identical
    values that should be merged into single cells in Excel. It returns a
    modified header row with duplicates replaced by empty strings, along with
    merge range information.

    Args:
        header_row (list): List of header values to process
        startrow (int): The Excel row number where this header will be written
        startcol (int): The Excel column number where the header starts

    Returns:
        tuple: A tuple containing:
            - header_final (list): Modified header row with duplicates as empty strings
            - merge_ranges (list): List of merge range tuples in format
                                  (row, start_col, end_col, value)

    Example:
        >>> get_unrepeated_header_row(['Sales', 'Sales', 'Profit'], 0, 0)
        (['Sales', '', 'Profit'], [(0, 0, 1, 'Sales')])

    Note:
        - Only consecutive duplicate values are merged
        - The first occurrence retains the value, subsequent duplicates become ''
        - Merge ranges are only created when there are 2+ consecutive duplicates
    """
    merge_ranges = []
    header_final = []
    i = 0
    mergestart = None
    mergevalue = ''

    # Compare each entry with the previous one
    # Prepend empty string to create entry_prev for first element
    for entry, entry_prev in zip(header_row, [''] + header_row[:-1]):
        if entry == entry_prev:
            # Duplicate found - add empty string to header
            header_final.append('')
        else:
            # New value encountered
            # If we were tracking a merge range, save it
            if (mergestart is not None) and (i + startcol - 1 > mergestart):
                merge_ranges.append((startrow, mergestart, startcol + i - 1, mergevalue))

            # Start tracking new potential merge range
            mergestart = startcol + i
            mergevalue = entry
            header_final.append(entry)

        i += 1

    # Handle final merge range if it exists
    if (mergestart is not None) and (i + startcol - 1 > mergestart):
        merge_ranges.append((startrow, mergestart, startcol + i - 1, mergevalue))

    return (header_final, merge_ranges)


def create_sheet(df: pd.DataFrame | list[pd.DataFrame], writer, sheet_name, **kwargs):
    if isinstance(df, pd.DataFrame):
        return create_sheet_single(df, writer, sheet_name, **kwargs)
    elif isinstance(df, list):
        if not df:
            return 1

        kwargs = kwargs.copy()
        current_startrow = kwargs.pop("startrow", 0)

        for df_i in df:
            if not isinstance(df_i, pd.DataFrame):
                raise TypeError("All items in the list passed to create_sheet must be pandas DataFrames.")

            table_kwargs = kwargs.copy()
            table_kwargs["startrow"] = current_startrow
            create_sheet_single(df_i, writer, sheet_name, **table_kwargs)
            current_startrow += df_i.shape[0] + 4
        return 1

    raise TypeError("create_sheet expects a pandas DataFrame or a list of pandas DataFrames.")

def create_sheet_single(df, writer, sheet_name, duplicate_header=False, hide_columns=[], startrow=0):
    """
    Create a professionally formatted Excel sheet from a pandas DataFrame.

    This function exports a DataFrame to Excel with enhanced formatting including:
    - Multi-level column and index support
    - Automatic cell merging for repeated headers
    - Frozen panes at header/index intersection
    - Auto-filters on all columns
    - Cell borders on all data
    - No gridlines for cleaner appearance
    - Centered and bolded header rows
    - Optional hidden columns in the final worksheet
    - Optional support for starting the exported table at a custom row

    Args:
        df (pandas.DataFrame): The DataFrame to export
        writer (pandas.ExcelWriter): An ExcelWriter object with engine='xlsxwriter'
        sheet_name (str): Name for the Excel sheet
        duplicate_header (bool, optional): If True, writes flattened column names
                                          below multi-level headers. Defaults to False.
        hide_columns (list, optional): A list of column labels to hide in the
                                      final Excel sheet. Defaults to [].
        startrow (int, optional): Starting row index for the DataFrame export.
                                 Defaults to 0.

    Returns:
        int: Returns 1 on success

    Example:
        >>> import pandas as pd
        >>> df = pd.DataFrame({'A': [1, 2], 'B': [3, 4]})
        >>> with pd.ExcelWriter('output.xlsx', engine='xlsxwriter') as writer:
        ...     create_sheet(df, writer, 'Sheet1', hide_columns=['B'], startrow=5)

    Note:
        - For single-level columns: writes DataFrame directly with minimal formatting
        - For multi-level columns: each level becomes a separate header row with merging
        - Index handling: preserves and formats both single and multi-level indexes
        - Auto-filter is applied to the last header row covering all data columns
    """
    if hide_columns is None:
        hide_columns = []
    startrow = int(startrow)

    # Create a copy to avoid modifying the original DataFrame
    b = df.copy()

    hidden_col_positions = []
    if hide_columns:
        for position, col_name in enumerate(b.columns):
            for hidden_name in hide_columns:
                if col_name == hidden_name or str(col_name) == str(hidden_name):
                    hidden_col_positions.append(position)
                    break

    # -------------------------------------------------------------------------
    # Determine index structure
    # -------------------------------------------------------------------------
    idx_names = list(b.index.names)

    # Check if DataFrame has a meaningful index
    if idx_names == [None]:
        # No named index (RangeIndex with no name)
        idx_nlevels = 0
    elif type(b.index) in [pd.Index, pd.RangeIndex]:
        # Single-level index
        idx_nlevels = 1
    else:
        # Multi-level index (MultiIndex)
        idx_nlevels = b.index.nlevels

    # -------------------------------------------------------------------------
    # Determine column structure
    # -------------------------------------------------------------------------
    col_nlevels = b.columns.nlevels
    col_values = b.columns.values

    # -------------------------------------------------------------------------
    # Analyze each data column for numeric/percentage formatting
    # -------------------------------------------------------------------------
    col_format_info = []
    for i in range(len(b.columns)):
        col = b.columns[i]
        series = b.iloc[:, i]
        col_name = '.'.join(str(c) for c in col) if isinstance(col, tuple) else col
        is_num = _is_numeric_col(series)
        is_pct = is_num and _is_pct_col_name(str(col_name))
        dec = _get_decimal_places(series) if is_num else 0
        dec_str = '.' + '0' * dec if dec > 0 else ''
        num_fmt = f'#,##0{dec_str}"%"' if is_pct else (f'#,##0{dec_str}' if is_num else None)
        col_format_info.append({
            'is_numeric': is_num,
            'is_pct': is_pct,
            'decimal_places': dec,
            'num_format': num_fmt,
        })

    # -------------------------------------------------------------------------
    # Calculate auto-filter range
    # -------------------------------------------------------------------------
    # These will be used later to set up auto-filter functionality
    merge_ranges_all = []
    autofilter_row_start = startrow + col_nlevels - 1  # Last header row
    autofilter_row_end = autofilter_row_start + b.shape[0]  # Last data row
    autofilter_col_end = b.shape[1] + idx_nlevels - 1  # Last column including index

    # -------------------------------------------------------------------------
    # Get workbook object for formatting
    # -------------------------------------------------------------------------
    workbook = writer.book

    # -------------------------------------------------------------------------
    # Write DataFrame based on column level complexity
    # -------------------------------------------------------------------------
    if col_nlevels == 1:
        # Simple case: write header-only empty DataFrame to create the worksheet
        write_b = b.reset_index() if idx_nlevels > 0 else b
        pd.DataFrame(columns=write_b.columns).to_excel(
            writer, sheet_name=sheet_name, index=False, startrow=startrow
        )
    else:
        # Complex case: multi-level columns
        header_start_row = startrow

        for col_level_i in range(col_nlevels):
            header_row = b.columns.get_level_values(col_level_i).tolist()
            header_row, merge_ranges = get_unrepeated_header_row(
                header_row, header_start_row + col_level_i, idx_nlevels
            )
            merge_ranges_all += merge_ranges
            write_df = pd.DataFrame(columns=header_row)
            write_df.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=header_start_row + col_level_i,
                index=False,
                startcol=idx_nlevels
            )

        last_header_row = header_start_row + col_nlevels - 1

        if idx_nlevels > 0:
            pd.DataFrame(columns=idx_names).to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=last_header_row,
                index=False
            )

        # Flatten column names and reset index
        b.columns = ['.'.join(str(x) for x in col).strip('.') for col in col_values]
        if idx_nlevels > 0:
            b = b.reset_index()

        write_b = b
        data_start_row = header_start_row + col_nlevels

    # -------------------------------------------------------------------------
    # Get worksheet (created by the header writes above)
    # -------------------------------------------------------------------------
    worksheet = writer.sheets[sheet_name]

    # Build xlsxwriter format objects for each data column
    col_xlsxfmts = []
    for info in col_format_info:
        if info['is_numeric']:
            col_xlsxfmts.append(workbook.add_format({'num_format': info['num_format']}))
        else:
            col_xlsxfmts.append(None)

    # Default format for datetime values (mirrors pandas to_excel behaviour)
    default_date_fmt = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss'})

    # Determine where data rows begin
    data_row_start = startrow + 1 if col_nlevels == 1 else data_start_row

    # For multi-level with duplicate_header: write flattened column names as extra row
    if col_nlevels > 1 and duplicate_header:
        for col_j, col_name in enumerate(write_b.columns):
            worksheet.write(data_row_start, col_j, str(col_name))
        data_row_start += 1
        autofilter_row_end += 1

    # Write data rows manually with column-specific number formats
    for row_i in range(len(write_b)):
        for col_j in range(write_b.shape[1]):
            val = write_b.iat[row_i, col_j]
            excel_row = data_row_start + row_i
            data_col_idx = col_j - idx_nlevels
            fmt = col_xlsxfmts[data_col_idx] if data_col_idx >= 0 else None
            if pd.isna(val):
                worksheet.write_blank(excel_row, col_j, None, fmt)
            elif isinstance(val, datetime.datetime) and fmt is None:
                # Apply a date format so Excel/openpyxl recognises the cell as a date
                worksheet.write(excel_row, col_j, val, default_date_fmt)
            else:
                worksheet.write(excel_row, col_j, val, fmt)

    # -------------------------------------------------------------------------
    # Apply worksheet-level formatting
    # -------------------------------------------------------------------------

    # -------------------------------------------------------------------------
    # Set column widths, number formats, and hidden status
    # -------------------------------------------------------------------------
    # Build header display strings for width computation (leaf level for multi-level)
    col_headers_for_width = []
    for j in range(idx_nlevels):
        name = idx_names[j]
        col_headers_for_width.append(str(name) if name is not None else '')
    for col in col_values:
        if isinstance(col, tuple):
            col_headers_for_width.append(str(col[-1]))
        else:
            col_headers_for_width.append(str(col))

    for col_j in range(write_b.shape[1]):
        header_len = len(col_headers_for_width[col_j])
        col_data = write_b.iloc[:, col_j]
        data_col_idx = col_j - idx_nlevels
        info = col_format_info[data_col_idx] if data_col_idx >= 0 else None
        fmt_obj = col_xlsxfmts[data_col_idx] if (info is not None and info['is_numeric']) else None

        max_data_len = 0
        for val in col_data:
            if pd.isna(val):
                continue
            if info and info['is_numeric']:
                try:
                    s = f"{float(val):,.{info['decimal_places']}f}"
                    if info['is_pct']:
                        s += '%'
                    max_data_len = max(max_data_len, len(s))
                except (TypeError, ValueError):
                    max_data_len = max(max_data_len, len(str(val)))
            else:
                max_data_len = max(max_data_len, len(str(val)))

        col_width = max(5, min(35, 5 + max(header_len, max_data_len)))
        is_hidden = (data_col_idx in hidden_col_positions) if data_col_idx >= 0 else False
        options = {'hidden': True} if is_hidden else {}
        worksheet.set_column(col_j, col_j, col_width, fmt_obj, options)

    # Freeze panes at the intersection of headers and index
    # This keeps headers and index visible when scrolling
    if startrow == 0:
        worksheet.freeze_panes(col_nlevels, idx_nlevels)

    # Remove gridlines for a cleaner appearance
    # 2 = hide both screen and print gridlines
    worksheet.hide_gridlines(2)

    # -------------------------------------------------------------------------
    # Merge cells for repeated header values
    # -------------------------------------------------------------------------
    # Convert merge ranges from row/col indices to Excel cell notation
    merge_ranges_all = convert_merge_index_to_range(merge_ranges_all)

    # Apply all merge ranges
    for cell_start, cell_end, cell_content in merge_ranges_all:
        worksheet.merge_range(f"{cell_start}:{cell_end}", cell_content)

    # -------------------------------------------------------------------------
    # Add auto-filter to data range
    # -------------------------------------------------------------------------
    # Auto-filter allows users to filter and sort data in Excel
    worksheet.autofilter(autofilter_row_start, 0, autofilter_row_end, autofilter_col_end)

    # -------------------------------------------------------------------------
    # Define and apply cell formats
    # -------------------------------------------------------------------------
    # Format for header rows: centered, bold text
    center = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'bold': True
    })

    # Additional number formats (currently commented out but available)
    # These can be applied to specific columns as needed
    # percent_format = workbook.add_format({'num_format': '0.00%'})
    # decimal_format = workbook.add_format({'num_format': '0.00'})
    # comma_format = workbook.add_format({'num_format': '#,##0'})

    # Apply centered format to all header rows
    row_start = startrow
    row_end = row_start + (autofilter_row_end - b.shape[0] + 1)
    for rownum in range(row_start, row_end):
        worksheet.set_row(rownum, None, center)

    # Example of how to apply number formats to specific columns:
    # worksheet.set_column('A:A', None, percent_format)
    # worksheet.set_column('B:B', None, decimal_format)
    # worksheet.set_column('C:C', None, comma_format)

    # -------------------------------------------------------------------------
    # Add borders to all cells
    # -------------------------------------------------------------------------
    # Define border format
    border_format = workbook.add_format({
        'border': 1,  # 1 = thin border
        'border_color': 'black'
    })

    # Apply borders using conditional formatting
    # This applies to the entire data range including headers, index, and data
    # Using a formula that always evaluates to TRUE ensures all cells get borders
    worksheet.conditional_format(
        startrow, 0,
        autofilter_row_end, autofilter_col_end,
        {
            'type': 'formula',
            'criteria': 'TRUE',
            'format': border_format
        }
    )

    return 1
