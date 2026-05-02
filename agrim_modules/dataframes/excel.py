"""
Excel export utilities for pandas DataFrames.

This module provides enhanced Excel export functionality for pandas DataFrames,
with special handling for multi-level column headers and indexes, automatic
cell merging, borders, and professional formatting.
"""

import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell


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


def create_sheet(df, writer, sheet_name, duplicate_header=False):
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

    Args:
        df (pandas.DataFrame): The DataFrame to export
        writer (pandas.ExcelWriter): An ExcelWriter object with engine='xlsxwriter'
        sheet_name (str): Name for the Excel sheet
        duplicate_header (bool, optional): If True, writes flattened column names
                                          below multi-level headers. Defaults to False.

    Returns:
        int: Returns 1 on success

    Example:
        >>> import pandas as pd
        >>> df = pd.DataFrame({'A': [1, 2], 'B': [3, 4]})
        >>> with pd.ExcelWriter('output.xlsx', engine='xlsxwriter') as writer:
        ...     create_sheet(df, writer, 'Sheet1')

    Note:
        - For single-level columns: writes DataFrame directly with minimal formatting
        - For multi-level columns: each level becomes a separate header row with merging
        - Index handling: preserves and formats both single and multi-level indexes
        - Auto-filter is applied to the last header row covering all data columns
    """
    # Create a copy to avoid modifying the original DataFrame
    b = df.copy()

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
    # Calculate auto-filter range
    # -------------------------------------------------------------------------
    # These will be used later to set up auto-filter functionality
    merge_ranges_all = []
    autofilter_row_start = col_nlevels - 1  # Last header row
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
        # Simple case: single-level columns
        # Just write the DataFrame directly
        if idx_nlevels == 0:
            # No index to write
            b.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            # Include index as regular column(s)
            b.reset_index().to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        # Complex case: multi-level columns
        # Need to write each header level separately and handle merging
        startrow = 0

        # Write each level of the column hierarchy
        for col_level_i in range(col_nlevels):
            # Get values for this level of the column hierarchy
            header_row = b.columns.get_level_values(col_level_i).tolist()

            # Process header row to identify merge ranges
            header_row, merge_ranges = get_unrepeated_header_row(
                header_row, startrow, idx_nlevels
            )

            # Accumulate all merge ranges
            merge_ranges_all += merge_ranges

            # Write this header row
            # Use empty DataFrame with these columns to write just the header
            write_df = pd.DataFrame(columns=header_row)
            write_df.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=startrow,
                index=False,
                startcol=idx_nlevels
            )

            startrow += 1

        # Write index names in the last header row if they exist
        if idx_nlevels > 0:
            pd.DataFrame(columns=idx_names).to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=startrow - 1,
                index=False
            )

        # Flatten column names for data section
        # Join multi-level column tuples with '.' and remove trailing/leading dots
        b.columns = ['.'.join(str(x) for x in col).strip('.') for col in col_values]

        # Reset index to include it as regular columns
        if idx_nlevels > 0:
            b = b.reset_index()

        # Write the actual data
        if duplicate_header:
            # Include the flattened column names as an additional header row
            autofilter_row_end += 1
            b.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=startrow,
                index=False
            )
        else:
            # Don't repeat column names (they're already in the multi-level headers)
            b.to_excel(
                writer,
                sheet_name=sheet_name,
                startrow=startrow,
                index=False,
                header=False
            )

    # -------------------------------------------------------------------------
    # Apply worksheet-level formatting
    # -------------------------------------------------------------------------
    worksheet = writer.sheets[sheet_name]

    # Freeze panes at the intersection of headers and index
    # This keeps headers and index visible when scrolling
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
    for rownum in range(autofilter_row_end - b.shape[0] + 1):
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
        0, 0,
        autofilter_row_end, autofilter_col_end,
        {
            'type': 'formula',
            'criteria': 'TRUE',
            'format': border_format
        }
    )

    return 1
