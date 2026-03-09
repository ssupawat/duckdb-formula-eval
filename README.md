# DuckDB Excel Formula Evaluator

A lightweight Excel formula evaluator library that converts Excel formulas to DuckDB SQL and executes them entirely within DuckDB. No Python evaluation - all computation happens in DuckDB.

## Features

- **Pure aggregates**: `SUM`, `AVERAGE`, `MAX`, `MIN`, `COUNT`, `COUNTIF`, `SUMIF`
- **Scalar arithmetic**: Basic math operations on cell references
- **IF statements**: Conditional formulas with nested conditions
- **Nested formulas**: Aggregates inside IF statements, IF with aggregate conditions
- **Cross-sheet VLOOKUP**: Lookup values across different sheets
- **Vectorized SQL evaluation**: Optimized columnar operations
- **Pure SQL architecture**: All evaluation happens within DuckDB

## Installation

**Recommended** - Install all dependencies from requirements.txt:

```bash
pip install -r requirements.txt
```

This installs: `duckdb`, `openpyxl`, `pandas`, `formulas`, `psutil`

**Manual installation:**

```bash
pip install duckdb openpyxl pandas formulas psutil
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

for sheet_name in excel_file.sheet_names:
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=0, engine='openpyxl')
    df.columns = [str(c).lower().replace(' ', '_') for c in df.columns]
    table_name = sheet_name.lower().replace(' ', '_')
    conn.register(table_name, df)

# Create evaluator
evaluator = FormulaEvaluator(conn)

# Apply formula to column (executes in DuckDB)
# Use Excel column syntax like D:D (column D) for column-wide operations
evaluator.apply_formula_to_column('=D:D*1.1', 'sheet1', 'bonus')

# Get results by querying DuckDB
result = conn.execute("SELECT SUM(bonus) FROM sheet1").fetchone()[0]
print(result)  # 825.0

# For debugging: see the SQL that would be generated
sql = evaluator.excel_to_sql('=SUM(D:D)', 'sheet1')
print(sql)  # SELECT (SELECT COALESCE(SUM("d"), 0) FROM sheet1)
```

**Note:** The library maps Excel column letters (A, B, C, D...) to actual column names in your data. Column `D` in your Excel file maps to the 4th column in your data.

### Multi-Step Pipeline Integration

```python
# Step 1: Apply formula to DuckDB table (using Excel column syntax)
evaluator.apply_formula_to_column('=D:D*1.1', 'sheet1', 'bonus')

# Step 2: Query results from DuckDB
total_bonus = conn.execute("SELECT SUM(bonus) FROM sheet1").fetchone()[0]

# Step 3: Data changes (from another step in the pipeline)
conn.execute("UPDATE sheet1 SET d = d * 2")

# Step 4: Recalculate formulas
evaluator.recalculate_all()

# Step 5: Query updated results
total_bonus = conn.execute("SELECT SUM(bonus) FROM sheet1").fetchone()[0]
```

## Supported Formula Types

### Pure Aggregates
```excel
=SUM(D:D)
=AVERAGE(D:D)
=MAX(D:D)
=MIN(D:D)
=COUNT(D:D)
=COUNTIF(C:C,"x")
=SUMIF(C:C,"x",D:D)
```

### Scalar Arithmetic (cell references)
```excel
=D1*1.07      # Multiply cell D1 by 1.07
=A1+B1+C1     # Add cells A1, B1, C1
=D:D*1.1      # Multiply entire column D by 1.1 (column operation)
```

### IF Statements
```excel
=IF(D1>80, D1*1.1, D1*0.9)           # Cell reference
=IF(D:D>80, D:D*1.1, D:D*0.9)        # Column operation (applies to all rows)
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

## API Reference

### `apply_formula_to_column(formula, sheet_name, target_column)`
Apply a formula to all rows in a target column.

**Parameters:**
- `formula` - Excel formula (use Excel syntax like `=D:D*1.1` for column D)
- `sheet_name` - Name of the sheet
- `target_column` - Name of column to store results

**Example:**
```python
evaluator.apply_formula_to_column('=D:D*0.1', 'sheet1', 'bonus')
# Applies 10% bonus to all values in column D
```

### `recalculate_all()`
Recalculate all stored formulas across all sheets.

**Example:**
```python
evaluator.recalculate_all()
```

### `excel_to_sql(formula, sheet_name, row_ctx=None)`
Convert an Excel formula to SQL without executing it. Useful for debugging.

**Parameters:**
- `formula` - Excel formula (e.g., "=SUM(D:D)")
- `sheet_name` - Name of the sheet
- `row_ctx` - Optional row context for cell references

**Returns:** SQL query string

**Example:**
```python
sql = evaluator.excel_to_sql('=SUM(D:D)', 'sheet1')
print(sql)  # SELECT (SELECT COALESCE(SUM("d"), 0) FROM sheet1)
```

### `get_formulas()`
Get all stored formulas.

**Returns:** Dictionary mapping sheet names to their stored formulas

**Example:**
```python
formulas = evaluator.get_formulas()
# {'sheet1': {'bonus': '=quantity*0.1'}}
```

### `last_sql` attribute
Contains the last generated SQL query (useful for debugging).

## Running Tests

```bash
# Run all 40 test cases
python3 test_formula_evaluator.py
```

## Implementation Details

### Excel to SQL Conversion Pipeline

**Conversion Order (Critical):**
1. VLOOKUP → SQL subqueries
2. Aggregates → SQL subqueries
3. IF → CASE expressions
4. Cell references → scalar values
5. Operators → SQL operators

**Input → Output Examples:**

| Input Formula | Generated SQL |
|--------------|---------------|
| `=IF(D1>80, D1*1.1, D1*0.9)` | `SELECT CASE WHEN 100.0 > 80 THEN 100.0 * 1.1 ELSE 100.0 * 0.9 END` |
| `=SUM(D:D)` | `SELECT (SELECT COALESCE(SUM("amount"), 0) FROM sheet1)` |
| `=VLOOKUP("A",Sheet2!A:B,2,0)` | `SELECT (SELECT COALESCE((SELECT label FROM sheet2 WHERE key = 'A' LIMIT 1), NULL))` |

## Project Structure

```
duckdb-formula-demo/
├── .gitignore
├── README.md
├── requirements.txt
├── formula_evaluator.py        # Library: FormulaEvaluator class
└── test_formula_evaluator.py   # Tests: 40 comprehensive test cases
```
