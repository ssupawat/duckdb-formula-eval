# DuckDB Excel Formula Evaluator

A lightweight Excel formula evaluator library using pure DuckDB SQL for multi-step pipeline integration.

## Features

- **Pure aggregates**: `SUM`, `AVERAGE`, `MAX`, `MIN`, `COUNT`, `COUNTIF`, `SUMIF`
- **Scalar arithmetic**: Basic math operations on cell references
- **IF statements**: Conditional formulas with nested conditions
- **Nested formulas**: Aggregates inside IF statements, IF with aggregate conditions
- **Cross-sheet VLOOKUP**: Lookup values across different sheets
- **Vectorized SQL evaluation**: 10-100x faster for simple arithmetic formulas
- **Pure SQL architecture**: All evaluation happens within DuckDB, no external dependencies

## Installation

```bash
pip install duckdb openpyxl pandas
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

# Evaluate formula
result = evaluator.evaluate_formula('=SUM(D:D)', 'sheet1')
print(result)  # 825.0

# With row context for cell references
result = evaluator.evaluate_formula('=IF(D1>80, D1*1.1, D1*0.9)', 'sheet1', {'D1': 100.0})
print(result)  # 110.0
```

### Multi-Step Pipeline Integration

```python
# Step 1: Apply formula to DuckDB table
evaluator.apply_formula_to_column('=B2*1.1', 'sheet1', 'bonus', context_column='quantity')

# Step 2: Data changes (from another step in the pipeline)
conn.execute("UPDATE sheet1 SET quantity = quantity * 2")

# Step 3: Recalculate formulas
evaluator.recalculate_all('sheet1')
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

## API Reference

### `evaluate_formula(formula, sheet_name, row_ctx=None)`
Evaluate an Excel formula.

**Parameters:**
- `formula` - Excel formula (e.g., "=SUM(D:D)")
- `sheet_name` - Name of the sheet
- `row_ctx` - Optional row context for cell references (e.g., {"D1": 100.0})

**Returns:** Formula result as float or string

### `apply_formula_to_column(formula, sheet_name, target_column, context_column=None)`
Apply a formula to all rows in a target column.

**Parameters:**
- `formula` - Excel formula
- `sheet_name` - Name of the sheet
- `target_column` - Name of column to store results
- `context_column` - Optional column for row context

**Example:**
```python
evaluator.apply_formula_to_column('=B2*0.1', 'sheet1', 'bonus', context_column='quantity')
```

### `recalculate_all(sheet_name=None)`
Recalculate all formulas for a sheet.

**Parameters:**
- `sheet_name` - Name of the sheet (if None, recalculate all sheets)

### `store_formula_at_cell(formula, sheet_name, row, col)`
Store a formula at a specific cell location.

**Parameters:**
- `formula` - Excel formula
- `sheet_name` - Name of the sheet
- `row` - Row number (1-indexed)
- `col` - Column letter (e.g., "A", "D")

**Example:**
```python
evaluator.store_formula_at_cell('=SUM(A:A)', 'sheet1', row=1, col='F')
```

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
duckdb-formula-eval/
├── .gitignore
├── README.md
├── requirements.txt
├── formula_evaluator.py        # Library: FormulaEvaluator class
└── test_formula_evaluator.py   # Tests: 40 comprehensive test cases
```
