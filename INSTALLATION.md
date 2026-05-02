# Installation Guide

## Install from GitHub (Recommended)

You can install this package directly from GitHub without cloning the repository:

### Basic Installation

```bash
pip install git+https://github.com/agrim-aggarwal/agrim_modules.git
```

### Install Specific Branch

```bash
pip install git+https://github.com/agrim-aggarwal/agrim_modules.git@branch-name
```

### Install Specific Tag/Version

```bash
pip install git+https://github.com/agrim-aggarwal/agrim_modules.git@v0.1.0
```

### Editable Installation (for development)

```bash
pip install -e git+https://github.com/agrim-aggarwal/agrim_modules.git#egg=agrim-modules
```

## Install from Local Clone

If you have cloned the repository:

### Using Poetry (Recommended for Development)

```bash
git clone https://github.com/agrim-aggarwal/agrim_modules.git
cd agrim_modules
poetry install
```

### Using pip

```bash
git clone https://github.com/agrim-aggarwal/agrim_modules.git
cd agrim_modules
pip install -e .
```

## Verify Installation

After installation, verify the package works:

```python
import agrim_modules
from agrim_modules import create_sheet

print("✓ agrim_modules installed successfully!")
```

## Requirements

- Python >= 3.11
- pandas >= 3.0.1
- numpy >= 2.4.3
- xlsxwriter >= 3.2.9

These dependencies will be automatically installed when you install the package.

## Usage Example

```python
import pandas as pd
from agrim_modules import create_sheet

# Create a simple DataFrame
df = pd.DataFrame({
    'Name': ['Alice', 'Bob', 'Charlie'],
    'Age': [25, 30, 35],
    'Salary': [50000, 60000, 70000]
})

# Export to Excel with professional formatting
with pd.ExcelWriter('output.xlsx', engine='xlsxwriter') as writer:
    create_sheet(df, writer, 'Employees')

print("✓ Excel file created successfully!")
```

## Troubleshooting

### Issue: "No module named 'agrim_modules'"

Solution: Make sure you've installed the package and are using the correct Python environment.

```bash
pip list | grep agrim
```

### Issue: Dependency conflicts

Solution: Create a fresh virtual environment:

```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install git+https://github.com/agrim-aggarwal/agrim_modules.git
```

### Issue: Poetry installation fails

Solution: Update Poetry and try again:

```bash
poetry self update
poetry install
```

## Uninstall

```bash
pip uninstall agrim-modules
```

## Support

For issues, please visit: https://github.com/agrim-aggarwal/agrim_modules/issues
