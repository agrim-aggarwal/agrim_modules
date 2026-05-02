"""
Test suite for agrim_modules.dataframes.excel functionality

This module contains comprehensive test cases for the create_sheet function,
demonstrating various DataFrame structures and export scenarios.

To run these tests:
    python tests/test_dataframes.py

Output:
    test_output.xlsx - Contains all test case outputs in separate sheets
"""

import pandas as pd
import numpy as np
import itertools
from datetime import datetime, timedelta
from agrim_modules import create_sheet


def create_sample_data():
    """
    Creates base sample data used across multiple test cases.

    Returns:
        pd.DataFrame: A DataFrame with categorical and numerical columns
    """
    np.random.seed(42)  # For reproducible results

    ser_a = [f"a{i}" for i in range(3)]
    ser_b = [f"b{i}" for i in range(2)]
    ser_c = [f"k{i}" for i in range(2)]

    df = pd.DataFrame(
        itertools.product(ser_a, ser_b, ser_c),
        columns=['c1', 'c2', 'c3']
    )
    df['c4'] = np.random.rand(len(df))
    df['c5'] = np.random.rand(len(df))

    return df


def test_simple_dataframe(writer):
    """
    Test Case 1: Simple DataFrame with no multi-level columns or index

    Scenario: Basic employee data table
    Expected: Standard table with auto-filter and frozen header row
    """
    print("Running Test 1: Simple DataFrame...")

    df = pd.DataFrame({
        'Name': ['Alice Johnson', 'Bob Smith', 'Charlie Brown', 'Diana Prince'],
        'Age': [25, 30, 35, 28],
        'Department': ['Engineering', 'Marketing', 'Engineering', 'Sales'],
        'Salary': [75000, 65000, 85000, 70000],
        'Years_Experience': [3, 5, 8, 4]
    })

    create_sheet(df, writer, '1_Simple_Table')
    print("  ✓ Created simple employee table")


def test_single_index_dataframe(writer):
    """
    Test Case 2: DataFrame with a single-level index

    Scenario: Regional sales data indexed by region
    Expected: Index column preserved, headers frozen
    """
    print("Running Test 2: Single Index DataFrame...")

    df = pd.DataFrame({
        'Q1_Sales': [100000, 150000, 120000, 90000],
        'Q2_Sales': [110000, 155000, 125000, 95000],
        'Q3_Sales': [105000, 160000, 130000, 92000],
        'Q4_Sales': [115000, 170000, 135000, 98000]
    }, index=['North', 'South', 'East', 'West'])

    df.index.name = 'Region'

    create_sheet(df, writer, '2_Single_Index')
    print("  ✓ Created regional sales with single index")


def test_multi_index_dataframe(writer):
    """
    Test Case 3: DataFrame with multi-level index

    Scenario: Sales data grouped by region and product category
    Expected: Both index levels preserved, proper freezing
    """
    print("Running Test 3: Multi-Index DataFrame...")

    index = pd.MultiIndex.from_product(
        [['North', 'South', 'East'], ['Electronics', 'Clothing']],
        names=['Region', 'Category']
    )

    df = pd.DataFrame({
        'Revenue': [50000, 30000, 60000, 35000, 55000, 32000],
        'Units_Sold': [500, 800, 600, 900, 550, 850],
        'Profit_Margin': [0.25, 0.35, 0.28, 0.38, 0.26, 0.36]
    }, index=index)

    create_sheet(df, writer, '3_Multi_Index')
    print("  ✓ Created multi-index sales data")


def test_simple_pivot_table(writer):
    """
    Test Case 4: Simple pivot table with 2-level columns

    Scenario: Pivot table showing min/max sales by region
    Expected: Merged cells for repeating column headers
    """
    print("Running Test 4: Simple Pivot Table...")

    base_data = create_sample_data()

    pivot = base_data.pivot_table(
        index=['c1'],
        columns=['c3'],
        values='c4',
        aggfunc=['min', 'max']
    )

    create_sheet(pivot, writer, '4_Simple_Pivot')
    print("  ✓ Created simple pivot with 2-level columns")


def test_complex_pivot_multi_index(writer):
    """
    Test Case 5: Complex pivot with multi-level index and columns

    Scenario: Sales aggregation by multiple dimensions
    Expected: Proper handling of both multi-index and multi-columns
    """
    print("Running Test 5: Complex Pivot (Multi-Index & Multi-Column)...")

    base_data = create_sample_data()

    pivot = base_data.pivot_table(
        index=['c1', 'c2'],
        columns=['c3'],
        values='c4',
        aggfunc=['min', 'max', 'mean']
    )

    create_sheet(pivot, writer, '5_Complex_Pivot_Index')
    print("  ✓ Created complex pivot with multi-level index")


def test_complex_pivot_multi_column(writer):
    """
    Test Case 6: Pivot with multi-level columns (3 levels)

    Scenario: Sales metrics broken down by multiple column dimensions
    Expected: Three levels of column headers with appropriate merging
    """
    print("Running Test 6: Complex Pivot (3-Level Columns)...")

    base_data = create_sample_data()

    pivot = base_data.pivot_table(
        index=['c1'],
        columns=['c3', 'c2'],
        values='c4',
        aggfunc=['min', 'max']
    )

    create_sheet(pivot, writer, '6_Complex_Pivot_Columns')
    print("  ✓ Created pivot with 3-level column hierarchy")


def test_duplicate_header_mode(writer):
    """
    Test Case 7: Using duplicate_header parameter

    Scenario: Pivot table with duplicate_header=True
    Expected: Both hierarchical headers and flattened column names visible
    """
    print("Running Test 7: Duplicate Header Mode...")

    base_data = create_sample_data()

    pivot = base_data.pivot_table(
        index=['c1'],
        values='c4',
        aggfunc=['min', 'max']
    )
    # Swap levels to create different structure
    pivot.columns = pivot.columns.swaplevel(0, 1)

    create_sheet(pivot, writer, '7_Duplicate_Header', duplicate_header=True)
    print("  ✓ Created pivot with duplicate headers shown")


def test_groupby_aggregation(writer):
    """
    Test Case 8: GroupBy aggregation with named aggregations

    Scenario: Grouped data with custom column names
    Expected: Clean output with multi-level index
    """
    print("Running Test 8: GroupBy Aggregation...")

    base_data = create_sample_data()

    grouped = base_data.groupby(['c1', 'c2']).agg(
        maximum=('c4', 'max'),
        minimum=('c4', 'min'),
        average=('c4', 'mean'),
        total=('c5', 'sum')
    )

    create_sheet(grouped, writer, '8_GroupBy_Agg')
    print("  ✓ Created grouped aggregation")


def test_single_column_groupby(writer):
    """
    Test Case 9: Simple groupby with single column output

    Scenario: Sum of values by single grouping column
    Expected: Single-level index, single-level column
    """
    print("Running Test 9: Single Column GroupBy...")

    base_data = create_sample_data()

    grouped = base_data.groupby(['c1'])[['c4']].sum()

    create_sheet(grouped, writer, '9_Single_Column_GroupBy')
    print("  ✓ Created single-column grouped data")


def test_datetime_data(writer):
    """
    Test Case 10: DataFrame with datetime index and values

    Scenario: Time series data with dates
    Expected: Proper date formatting in Excel
    """
    print("Running Test 10: DateTime Data...")

    dates = pd.date_range('2024-01-01', periods=10, freq='D')
    df = pd.DataFrame({
        'Temperature': np.random.randint(60, 90, 10),
        'Humidity': np.random.randint(30, 70, 10),
        'Precipitation': np.random.rand(10) * 2
    }, index=dates)

    df.index.name = 'Date'

    create_sheet(df, writer, '10_DateTime_Series')
    print("  ✓ Created time series data")


def test_wide_pivot_table(writer):
    """
    Test Case 11: Wide pivot table with many columns

    Scenario: Product performance across multiple metrics and categories
    Expected: Handles wide tables with many merged header cells
    """
    print("Running Test 11: Wide Pivot Table...")

    np.random.seed(42)

    # Create sample product data
    products = ['Product_A', 'Product_B', 'Product_C']
    regions = ['North', 'South', 'East', 'West']
    months = ['Jan', 'Feb', 'Mar']

    data = []
    for product in products:
        for region in regions:
            for month in months:
                data.append({
                    'Product': product,
                    'Region': region,
                    'Month': month,
                    'Revenue': np.random.randint(10000, 50000),
                    'Units': np.random.randint(100, 500)
                })

    df = pd.DataFrame(data)

    pivot = df.pivot_table(
        index='Product',
        columns=['Region', 'Month'],
        values=['Revenue', 'Units'],
        aggfunc='sum'
    )

    create_sheet(pivot, writer, '11_Wide_Pivot')
    print("  ✓ Created wide pivot table")


def test_percentage_data(writer):
    """
    Test Case 12: DataFrame with percentage values

    Scenario: Performance metrics as percentages
    Expected: Values displayed as decimals (formatting can be customized)
    """
    print("Running Test 12: Percentage Data...")

    df = pd.DataFrame({
        'Department': ['Sales', 'Marketing', 'Engineering', 'HR'],
        'Growth_Rate': [0.15, 0.08, 0.22, 0.05],
        'Turnover_Rate': [0.12, 0.07, 0.04, 0.09],
        'Budget_Utilization': [0.95, 0.88, 0.92, 0.78]
    })

    df = df.set_index('Department')

    create_sheet(df, writer, '12_Percentage_Data')
    print("  ✓ Created percentage data table")


def test_mixed_aggregations(writer):
    """
    Test Case 13: Pivot with different aggregations per column

    Scenario: Different statistics for different value columns
    Expected: Complex multi-level column structure
    """
    print("Running Test 13: Mixed Aggregations...")

    np.random.seed(42)

    df = pd.DataFrame({
        'Category': ['A', 'B', 'A', 'B', 'A', 'B'] * 3,
        'Subcategory': ['X', 'X', 'Y', 'Y', 'Z', 'Z'] * 3,
        'Sales': np.random.randint(1000, 5000, 18),
        'Costs': np.random.randint(500, 3000, 18),
        'Quantity': np.random.randint(10, 100, 18)
    })

    pivot = df.pivot_table(
        index='Category',
        columns='Subcategory',
        values=['Sales', 'Costs', 'Quantity'],
        aggfunc={'Sales': ['sum', 'mean'], 'Costs': 'sum', 'Quantity': 'sum'}
    )

    create_sheet(pivot, writer, '13_Mixed_Aggregations')
    print("  ✓ Created mixed aggregations pivot")


def test_no_index_dataframe(writer):
    """
    Test Case 14: DataFrame with RangeIndex (no named index)

    Scenario: Simple data without meaningful index
    Expected: No index column in output, just data columns
    """
    print("Running Test 14: No Index DataFrame...")

    df = pd.DataFrame({
        'Transaction_ID': [f'TXN{i:04d}' for i in range(1, 21)],
        'Amount': np.random.randint(10, 1000, 20),
        'Status': np.random.choice(['Completed', 'Pending', 'Failed'], 20)
    })

    create_sheet(df, writer, '14_No_Index')
    print("  ✓ Created DataFrame without named index")


def test_empty_dataframe(writer):
    """
    Test Case 15: Edge case - Empty DataFrame

    Scenario: DataFrame with columns but no data
    Expected: Headers displayed, no data rows
    """
    print("Running Test 15: Empty DataFrame...")

    df = pd.DataFrame(columns=['Column_A', 'Column_B', 'Column_C'])

    create_sheet(df, writer, '15_Empty_DataFrame')
    print("  ✓ Created empty DataFrame")


def test_single_row_dataframe(writer):
    """
    Test Case 16: Edge case - Single row DataFrame

    Scenario: DataFrame with only one data row
    Expected: Proper formatting with auto-filter
    """
    print("Running Test 16: Single Row DataFrame...")

    df = pd.DataFrame({
        'Metric': ['Total Revenue'],
        'Value': [1250000],
        'Target': [1000000],
        'Variance': [0.25]
    })

    create_sheet(df, writer, '16_Single_Row')
    print("  ✓ Created single-row DataFrame")


def test_single_column_dataframe(writer):
    """
    Test Case 17: Edge case - Single column DataFrame

    Scenario: DataFrame with only one column
    Expected: Proper formatting with auto-filter
    """
    print("Running Test 17: Single Column DataFrame...")

    df = pd.DataFrame({
        'Values': [10, 20, 30, 40, 50]
    })
    df.index.name = 'Index'

    create_sheet(df, writer, '17_Single_Column')
    print("  ✓ Created single-column DataFrame")


def test_large_numbers(writer):
    """
    Test Case 18: DataFrame with large numbers

    Scenario: Financial data with millions/billions
    Expected: Numbers displayed fully (custom formatting possible)
    """
    print("Running Test 18: Large Numbers...")

    df = pd.DataFrame({
        'Company': ['TechCorp', 'FinanceInc', 'RetailCo', 'EnergyLtd'],
        'Revenue_USD': [5_234_567_890, 12_456_789_012, 3_456_789_123, 8_901_234_567],
        'Employees': [45_000, 78_000, 120_000, 34_000],
        'Market_Cap': [150_000_000_000, 340_000_000_000, 89_000_000_000, 210_000_000_000]
    })

    df = df.set_index('Company')

    create_sheet(df, writer, '18_Large_Numbers')
    print("  ✓ Created large numbers table")


def run_all_tests():
    """
    Main test runner that executes all test cases and generates output Excel file.
    """
    output_file = 'test_output.xlsx'

    print("=" * 70)
    print("Starting agrim_modules Test Suite")
    print("=" * 70)
    print(f"Output file: {output_file}")
    print()

    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # Run all test cases
        test_simple_dataframe(writer)
        test_single_index_dataframe(writer)
        test_multi_index_dataframe(writer)
        test_simple_pivot_table(writer)
        test_complex_pivot_multi_index(writer)
        test_complex_pivot_multi_column(writer)
        test_duplicate_header_mode(writer)
        test_groupby_aggregation(writer)
        test_single_column_groupby(writer)
        test_datetime_data(writer)
        test_wide_pivot_table(writer)
        test_percentage_data(writer)
        test_mixed_aggregations(writer)
        test_no_index_dataframe(writer)
        test_empty_dataframe(writer)
        test_single_row_dataframe(writer)
        test_single_column_dataframe(writer)
        test_large_numbers(writer)

    print()
    print("=" * 70)
    print(f"✓ All tests completed successfully!")
    print(f"✓ Output saved to: {output_file}")
    print(f"✓ Total test cases: 18")
    print("=" * 70)
    print()
    print("Open the Excel file to review all test case outputs.")


if __name__ == '__main__':
    run_all_tests()            