# agrim_modules

A Python utility library for Data Science, Analytics, and Engineering projects, providing enhanced tools for working with pandas DataFrames.

## Overview

This repository contains custom-developed modules that extend pandas functionality, particularly for creating professionally formatted Excel exports with multi-level headers, merged cells, and automatic formatting.

## Features

### DataFrame to Excel Conversion

The core functionality revolves around the `create_sheet` function, which converts pandas DataFrames to Excel with advanced formatting:

- ✅ **Multi-level column headers**: Properly handles DataFrames with hierarchical column structures
- ✅ **Multi-level index support**: Works with single or multi-level row indexes
- ✅ **Automatic cell merging**: Merges duplicate values in multi-level headers for cleaner appearance
- ✅ **Freeze panes**: Automatically freezes header rows and index columns for easier navigation
- ✅ **Auto-filters**: Adds Excel auto-filter functionality to all data columns
- ✅ **Professional formatting**: Centers and bolds header rows automatically

## Installation

### From Source

```bash
git clone https://github.com/yourusername/agrim_modules.git
cd agrim_modules
pip install -e .
```

### Requirements

- Python >= 3.11
- pandas >= 3.0.1
- numpy >= 2.4.3
- xlsxwriter >= 3.2.9

## Usage

### Basic Example

```python
import pandas as pd
from agrim_modules import create_sheet

# Create a simple DataFrame
df = pd.DataFrame({
    'Name': ['Alice', 'Bob', 'Charlie'],
    'Age': [25, 30, 35],
    'Salary': [50000, 60000, 70000]
})

# Write to Excel with formatting
with pd.ExcelWriter('output.xlsx', engine='xlsxwriter') as writer:
    create_sheet(df, writer, 'Employees')
```

### Advanced Example: Multi-level Columns

```python
import pandas as pd
import numpy as np
from agrim_modules import create_sheet

# Create DataFrame with multi-level columns (e.g., from pivot_table)
data = {
    ('Sales', 'Q1', 'Revenue'): [100, 200, 300],
    ('Sales', 'Q1', 'Units'): [10, 20, 30],
    ('Sales', 'Q2', 'Revenue'): [150, 250, 350],
    ('Sales', 'Q2', 'Units'): [15, 25, 35],
}
df = pd.DataFrame(data, index=['North', 'South', 'East'])
df.index.name = 'Region'

# Export to Excel
with pd.ExcelWriter('sales_report.xlsx', engine='xlsxwriter') as writer:
    create_sheet(df, writer, 'Sales Data')
```

### Example: Pivot Tables

```python
import pandas as pd
from agrim_modules import create_sheet

# Sample data
df = pd.DataFrame({
    'Region': ['North', 'South', 'North', 'South'],
    'Product': ['A', 'A', 'B', 'B'],
    'Sales': [100, 150, 200, 250],
    'Quantity': [10, 15, 20, 25]
})

# Create pivot table
pivot = df.pivot_table(
    index=['Region'],
    columns=['Product'],
    values='Sales',
    aggfunc=['sum', 'mean']
)

# Export with proper formatting
with pd.ExcelWriter('pivot_report.xlsx', engine='xlsxwriter') as writer:
    create_sheet(pivot, writer, 'Pivot Analysis')
```

### Example: Multiple DataFrames

```python
import pandas as pd
from agrim_modules import create_sheet

df1 = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6]})
df2 = pd.DataFrame({'X': [7, 8, 9], 'Y': [10, 11, 12]})

# Write multiple sheets to one Excel file
with pd.ExcelWriter('multi_sheet.xlsx', engine='xlsxwriter') as writer:
    create_sheet(df1, writer, 'First Sheet')
    create_sheet(df2, writer, 'Second Sheet')
```

## API Reference

### `create_sheet(df, writer, sheet_name, duplicate_header=False)`

Creates a formatted Excel sheet from a pandas DataFrame.

#### Parameters

- **`df`** (`pandas.DataFrame`): The DataFrame to export
- **`writer`** (`pandas.ExcelWriter`): An ExcelWriter object with engine='xlsxwriter'
- **`sheet_name`** (`str`): Name for the Excel sheet
- **`duplicate_header`** (`bool`, optional):
  - Default: `False`
  - When `True`, writes the flattened column names below multi-level headers
  - Useful when you want to see both the hierarchical structure and the full column names

#### Returns

- `int`: Returns `1` on success

#### Behavior

1. **Single-level columns**: Writes DataFrame directly with minimal formatting
2. **Multi-level columns**:
   - Writes each level of the column hierarchy as separate header rows
   - Merges cells where values repeat horizontally
   - Flattens column names for the data section
3. **Index handling**:
   - Preserves and formats single or multi-level indexes
   - Index names are written in the header row
4. **Formatting applied**:
   - Freeze panes at the intersection of headers and indexes
   - Auto-filter on all data columns
   - Center-aligned, bold header rows

## Example Output

The function creates Excel files that look like this:

### Before (standard pandas.to_excel):
```
| | k0_min | k0_max | k1_min | k1_max |
|a0| 0.123 | 0.456 | 0.789 | 0.321 |
```

### After (create_sheet):
```
|   |    k0     |    k1     |
|   | min | max | min | max |
|a0 |0.123|0.456|0.789|0.321|
```
(with merged cells for "k0" and "k1", frozen panes, and auto-filters)

## Use Cases

This library is particularly useful for:

- **Reporting**: Creating professional Excel reports from pandas analysis
- **Pivot table exports**: Properly formatting complex pivot tables
- **Data delivery**: Providing clean, user-friendly data exports to stakeholders
- **Dashboard data**: Exporting multi-dimensional data with clear hierarchical structure

## Testing

Run the test suite to see examples of various DataFrame structures:

```bash
python tests/test_dataframes.py
```

This will generate a `test.xlsx` file with examples of:
- Simple DataFrames
- Single-level pivot tables
- Multi-level pivot tables
- Grouped aggregations
- DataFrames with and without indexes

## Contributing

Contributions are welcome! This library is designed to grow with commonly needed utilities for data science work.

## License

[Add your license here]

## Author

Agrim Aggarwal

## Roadmap

Future enhancements may include:
- Column width auto-sizing
- Custom number formatting per column
- Conditional formatting support
- Chart generation
- Additional export formats
- More data transformation utilities
