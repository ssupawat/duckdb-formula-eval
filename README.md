# DuckDB Excel Formula Evaluation

A high-performance Excel formula evaluator using DuckDB for bulk SQL operations.

## Background

This project demonstrates using DuckDB for Excel formula evaluation, achieving **2-5x slower than JavaScript** (SheetJS) compared to the original **65-970x slower** implementation.

## Performance Results

### Standard Tests (Single Sheet, =A{i}+B{i} formulas)

| Rows | DuckDB Time | DuckDB Peak | JS Time | LO Time | Speedup vs JS |
|------|-------------|-------------|---------|---------|---------------|
| 10K  | 0.339s      | 115 MB      | 0.14s   | 1.01s   | 2.4x slower   |
| 50K  | 1.628s      | 168 MB      | 0.46s   | 0.88s   | 3.5x slower   |
| 100K | 3.306s      | 234 MB      | 0.91s   | 1.45s   | 3.6x slower   |
| 200K | 6.733s      | 359 MB      | 1.93s   | 2.04s   | 3.5x slower   |

### Complex Formula Benchmarks (IF statements with numexpr)

| Formula Type | Example | Rows | Time | Throughput |
|--------------|---------|------|------|------------|
| **Complex** | `=IF(A{i}="x", B{i}*1.1, B{i}*0.9)` | 10K | 1.67s | ~6K rows/s |
| **Complex** | `=IF(A{i}="x", B{i}*1.1, B{i}*0.9)` | 50K | 8.32s | ~6K rows/s |
| **Complex** | `=IF(A{i}="x", B{i}*1.1, B{i}*0.9)` | 100K | 17.54s | ~6K rows/s |

**Note:** Complex formulas require per-row scalar evaluation (numexpr), making them ~4x slower than simple formulas processed entirely in SQL.

See [BENCHMARK_COMPARISON.md](BENCHMARK_COMPARISON.md) for detailed performance analysis.

### Performance Improvement Over Original

| Rows | Original Time | Optimized Time | Speedup |
|------|---------------|----------------|---------|
| 10K  | 9.09s         | 0.339s         | **26.8x** |
| 50K  | 215.73s       | 1.628s         | **132.5x** |
| 100K | 883.58s       | 3.306s         | **267.3x** |

## Files

| File | Description |
|------|-------------|
| `measure_duckdb_optimized.py` | Optimized implementation using pure SQL |
| `test_formula_evaluator.py` | Comprehensive evaluator (aggregates, IF, VLOOKUP) |
| `measure_complex_formulas.py` | Complex formula benchmark (IF, VLOOKUP, nested) |
| `generate_test_files.py` | Generate simple test Excel files |
| `generate_complex_test_files.py` | Generate complex formula test files |
| `run_benchmark.sh` | Run simple formula benchmark suite |
| `run_complex_benchmark.sh` | Run complex formula benchmark suite |
| `run_comparison.sh` | Compare simple vs complex formula performance |
| `COMPARISON.md` | Detailed comparison and results |
| `BENCHMARK_COMPARISON.md` | Simple vs complex formula performance analysis |

## Usage

### Run Optimized Benchmark

```bash
python3 measure_duckdb_optimized.py test_files/test_10000.xlsx
```

### Run Full Benchmark

```bash
bash run_benchmark.sh
```

### Run Complex Formula Benchmark

```bash
# Generate complex test files
python3 generate_complex_test_files.py

# Run complex formula benchmark
python3 measure_complex_formulas.py test_files_complex/complex_10000.xlsx

# Run all complex formula benchmarks
bash run_complex_benchmark.sh

# Compare simple vs complex formulas
bash run_comparison.sh
```

### Run Comprehensive Formula Tests (14 test cases)

```bash
python3 test_formula_evaluator.py
```

**Supported formulas:**
- Pure aggregates: `SUM(D:D)`, `AVERAGE(D:D)`, `MAX(D:D)`
- Conditional aggregates: `SUMIF(C:C,"x",D:D)`, `COUNTIF(C:C,"x")`
- Scalar arithmetic: `=D1*1.07`
- IF statements: `=IF(D1>80, D1*1.1, D1*0.9)`
- Nested formulas: `=IF(SUMIF(C:C,"x",D:D)>100, D1*1.07, 0)`
- Cross-sheet VLOOKUP: `=VLOOKUP(A1,Sheet2!A:B,2,0)`

## Implementation Details

### Key Optimization

**Before (Slow)**: Per-cell processing
```
openpyxl → Python loop → formulas lib → DuckDB queries → openpyxl write
```

**After (Fast)**: Pure SQL bulk operations
```
Excel → DuckDB (bulk load) → SQL queries (bulk) → DuckDB → Excel
```

### Formula Pattern Mapping

| Excel Pattern | SQL Equivalent |
|---------------|----------------|
| `=A{i}+B{i}` | `SELECT "_row", "a", "b", "a" + "b" AS "c" FROM sheet1` |
| `=Sheet1!A{i}` | `SELECT t1._row, t1."doubled", t2."value" AS "from_sheet1" FROM sheet2 t1 JOIN sheet1 t2 ON t1._row = t2._row` |
| `=A{i}*2` | `SELECT "_row", "from_sheet1", "from_sheet1" * 2 AS "doubled" FROM sheet2` |

### Two-Phase Formula Evaluation (test_formula_evaluator.py)

1. **Phase 1 (DuckDB SQL)**: Extract and compute all aggregates
2. **Phase 2 (numexpr)**: Substitute aggregates and evaluate scalar expressions safely and efficiently

Using **numexpr** instead of Python's `eval()` provides:
- **2-10x faster** scalar expression evaluation (C-optimized)
- **Security**: No arbitrary code execution risk
- **Compatibility**: Works with numpy arrays and Python variables

This hybrid approach supports complex nested formulas like:
```excel
=IF(SUMIF(C:C,"x",D:D)/COUNTIF(C:C,"x")>50, D1*2, D1)
```

## Requirements

```
duckdb>=0.9.0
pandas>=2.0.0
openpyxl>=3.0.0
formulas>=1.0.0
psutil>=5.9.0
numexpr>=2.8.0
```

## Installation

```bash
pip install duckdb openpyxl pandas formulas psutil numexpr
```

## License

MIT
