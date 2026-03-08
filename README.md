# DuckDB Excel Formula Evaluator

A Excel formula evaluator using DuckDB for SQL-based evaluation.

## Features

- **Pure aggregates**: `SUM`, `AVERAGE`, `MAX`, `MIN`, `COUNTIF`, `SUMIF`
- **Scalar arithmetic**: Basic math operations on cell references
- **IF statements**: Conditional formulas with nested conditions
- **Nested formulas**: Aggregates inside IF statements, IF with aggregate conditions
- **Cross-sheet VLOOKUP**: Lookup values across different sheets
- **Vectorized SQL evaluation**: 10-100x faster for simple arithmetic formulas
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
# Run all 48 test cases
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

## Core Concepts

The DuckDB Formula Demo implements three core concepts for high-performance formula evaluation:

### 1. Pattern Detection

**Where:** `formula_evaluator.py:37-80` → `_parse_formula_pattern()`

**Purpose:** Classify formula to determine evaluation strategy

**Input → Output Examples:**

| Input Formula | Detected Pattern | Type | Next Step |
|--------------|------------------|------|-----------|
| `=A2+B2` | `A + B` | simple | Vectorized SQL |
| `=A2*2` | `A * 2` | scalar | Vectorized SQL |
| `=Sheet1!A2` | source=Sheet1, col=A | cross_sheet | Vectorized SQL |
| `=SUM(D:D)` | (no match) | complex | Two-Phase |
| `=IF(D1>80, D1*1.1, D1)` | (no match) | complex | Two-Phase |

**Key Regex Patterns:**
```python
r'^([A-Z])\d+\s*([+\-*/])\s*([A-Z])\d+$'  # A2+B2
r'^([A-Z])\d+\s*([+\-*/])\s*(\d+(?:\.\d+)?)$'  # A2*2
r'^([A-Za-z0-9_]+)!([A-Z])\d+$'  # Sheet1!A2
```

---

### 2. Vectorized SQL Evaluation

**Where:** `formula_evaluator.py:82-131` → `_evaluate_vectorized()`

**Purpose:** Evaluate simple formulas on entire columns using single SQL query

**Input → Processing → Output:**

**Input:**
- Formula: `=A2+B2`
- Sheet data (5 rows × 3 columns):
```
     A    B    C
1   10   20   30
2   15   25   35
3   20   30   40
4   25   35   45
5   30   40   50
```

**Processing:**
```python
# 1. Build column map
col_map = {'A': 'col0', 'B': 'col1', 'C': 'col2'}

# 2. Parse pattern: "A + B"
# 3. Generate SQL
sql = 'SELECT "col0" + "col1" FROM sheet1'

# 4. Execute
result = conn.execute(sql).fetchdf()
```

**Output:** `pd.Series([30, 40, 50, 60, 70])`

**Performance:** 10K rows in ~0.014s (vs ~17s per-cell)

---

### 3. Two-Phase Decomposition

**Where:** `formula_evaluator.py:150-364`
- Phase 1: `_resolve_aggregates()` (lines 150-181)
- Phase 2: `_evaluate_scalar()` (lines 306-364)

**Purpose:** Handle complex formulas with aggregates, IF, VLOOKUP

**Input → Processing → Output:**

**Input:**
- Formula: `=IF(SUM(D:D)>100, D1*1.1, D1*0.9)`
- Sheet data: Column D = [100, 200, 150, 75, 300]
- Row context: `{"D1": 100.0}`

**Phase 1: Resolve Aggregates**
```
Input:  =IF(SUM(D:D)>100, D1*1.1, D1*0.9)
        ↓
SQL:    SELECT COALESCE(SUM("col3"), 0) FROM sheet1
        ↓
Result: 825.0
        ↓
Output: =IF(825.0>100, D1*1.1, D1*0.9)
```

**Phase 2: Evaluate Scalar**
```
Input:  =IF(825.0>100, D1*1.1, D1*0.9)
        ↓
Substitute D1=100: =IF(825.0>100, 100*1.1, 100*0.9)
        ↓
Evaluate condition (825.0>100 = TRUE): 100*1.1
        ↓
numexpr: where(825.0>100, 100*1.1, 100*0.9)
        ↓
Output: 110.0
```

**Why Two Phases?**
- Aggregates need SQL (operate on entire columns)
- Cell references need scalar values (from row_ctx)
- Separation allows each phase to use optimal tool

## Implementation Details

### Vectorized SQL Evaluation

The evaluator includes a **pattern detection + vectorized evaluation** optimization for simple formulas:

| Pattern Type | Example | Detection | Evaluation Method |
|-------------|---------|-----------|-------------------|
| Simple arithmetic | `=A2+B2`, `=D2*E2` | Two-column arithmetic | SQL on entire column |
| Scalar operation | `=A2*2`, `=B2/10` | Column + constant | SQL on entire column |
| Cross-sheet | `=Sheet1!A2` | Sheet reference | SQL from other sheet |
| Complex | `=SUM(D:D)`, `IF(...)` | Aggregates/IF/VLOOKUP | Two-phase (aggregate + numexpr) |

**Performance comparison for 10K rows:**
- Original (per-cell): ~17 seconds
- Optimized (vectorized): ~0.014 seconds
- **Speedup: 1200x**

The optimization is transparent - existing code using `evaluate_formula()` continues to work without changes.

### Two-Phase Formula Evaluation

For complex formulas (aggregates, IF statements, VLOOKUP), the evaluator uses a two-phase approach:

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
│   ├── evaluate_formula()      # Main entry point (auto-detects vectorized vs two-phase)
│   ├── _parse_formula_pattern() # Detects simple patterns for vectorized evaluation
│   ├── _evaluate_vectorized()   # Vectorized SQL evaluation for simple formulas
│   ├── _resolve_aggregates()    # Phase 1: Compute aggregates via SQL
│   └── _evaluate_scalar()       # Phase 2: Evaluate scalar expressions via numexpr
├── test_formula_evaluator.py   # Tests: 48 comprehensive test cases
├── generate_test_files.py      # Test data generator
└── test_files/
    ├── simple_10k.xlsx         # Simple formula test file
    └── complex_10k.xlsx        # Complex formula test file
```
