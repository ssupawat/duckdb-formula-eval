# DuckDB Excel Formula Evaluator

A Excel formula evaluator using DuckDB for SQL-based evaluation.

## Features

- **Pure aggregates**: `SUM`, `AVERAGE`, `MAX`, `MIN`, `COUNTIF`, `SUMIF`
- **Scalar arithmetic**: Basic math operations on cell references
- **IF statements**: Conditional formulas with nested conditions
- **Nested formulas**: Aggregates inside IF statements, IF with aggregate conditions
- **Cross-sheet VLOOKUP**: Lookup values across different sheets
- **Hybrid evaluation**: SQL for aggregates + numexpr for scalar expressions

## Installation

```bash
pip install duckdb openpyxl pandas numexpr
```

Or using requirements.txt:

```bash
pip install -r requirements.txt
```

## Usage

### Basic Library Usage

```python
from formula_evaluator import FormulaEvaluator
import duckdb
import pandas as pd

# Load Excel file
excel_file = pd.ExcelFile('input.xlsx', engine='openpyxl')

# Create DuckDB connection and load data
conn = duckdb.connect(':memory:')
sheets_data = {}

for sheet_name in excel_file.sheet_names:
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=0, engine='openpyxl')
    df.columns = [str(c).lower().replace(' ', '_') for c in df.columns]
    table_name = sheet_name.lower().replace(' ', '_')
    sheets_data[table_name] = df
    conn.register(table_name, df)

# Create evaluator
evaluator = FormulaEvaluator(conn, sheets_data)

# Evaluate formulas
result = evaluator.evaluate_formula('=SUM(D:D)', 'sheet1')
print(result)  # 825.0

# With row context for cell references
result = evaluator.evaluate_formula('=IF(D1>80, D1*1.1, D1*0.9)', 'sheet1', {'D1': 100.0})
print(result)  # 110.0
```

### Running Tests

```bash
# Run all 14 test cases
python3 test_formula_evaluator.py
```

### Generating Test Files

```bash
# Generate simple 10K test file
python3 generate_test_files.py 10000

# Generate all simple test files
python3 generate_test_files.py

# Generate complex 10K test file
python3 generate_test_files.py --complex 10000

# Generate all complex test files
python3 generate_test_files.py --complex
```

## Supported Formula Types

### Pure Aggregates
```excel
=SUM(D:D)
=AVERAGE(D:D)
=MAX(D:D)
=MIN(D:D)
=COUNTIF(C:C,"x")
=SUMIF(C:C,"x",D:D)
```

### Scalar Arithmetic
```excel
=D1*1.07
=A1+B1+C1
```

### IF Statements
```excel
=IF(D1>80, D1*1.1, D1*0.9)
```

### Nested Formulas
```excel
=IF(SUMIF(C:C,"x",D:D)>100, D1*1.07, 0)
=IF(SUMIF(C:C,"x",D:D)/COUNTIF(C:C,"x")>50, D1*2, D1)
```

### Arithmetic on Aggregates
```excel
=SUM(D:D)*0.1
=AVERAGE(D:D)*1.2
```

### Cross-sheet VLOOKUP
```excel
=VLOOKUP(A1,Sheet2!A:B,2,0)
```

## Implementation Details

### Two-Phase Formula Evaluation

1. **Phase 1 (DuckDB SQL)**: Extract and compute all aggregate functions
2. **Phase 2 (numexpr)**: Substitute aggregates and evaluate scalar expressions safely

Using **numexpr** instead of Python's `eval()` provides:
- **2-10x faster** scalar expression evaluation (C-optimized)
- **Security**: No arbitrary code execution risk
- **Compatibility**: Works with numpy arrays and Python variables

## Project Structure

```
duckdb-formula-demo/
├── .gitignore
├── README.md
├── requirements.txt
├── formula_evaluator.py        # Library: FormulaEvaluator class
├── test_formula_evaluator.py   # Tests: imports and tests FormulaEvaluator
├── generate_test_files.py      # Test data generator
└── test_files/
    ├── simple_10k.xlsx         # Simple formula test file
    └── complex_10k.xlsx        # Complex formula test file
```
