# Formula Evaluation Benchmark Comparison

This document compares the performance of different formula evaluation approaches using the DuckDB + numexpr implementation.

## Benchmark Results Summary

### 1. Simple vs Complex Formulas

| Formula Type | Example | Approach | 10K Time | 50K Time | 100K Time | Throughput |
|-------------|---------|----------|----------|----------|-----------|------------|
| **Simple** | `=A{i}+B{i}` | Pure SQL | 0.42s | 1.68s | 4.15s | ~24K-30K rows/s |
| **Complex** | `=IF(A{i}="x", B{i}*1.1, B{i}*0.9)` | SQL + numexpr | 1.67s | 8.32s | 17.54s | ~6K rows/s |
| **Slowdown** | - | - | **3.99x** | **4.95x** | **4.23x** | - |

### 2. numexpr vs eval() (10K formulas)

| Method | Time | Throughput | Speedup |
|--------|------|------------|--------|
| **numexpr** | 0.076s | 131K/s | 1.00x (baseline) |
| **eval()** | 0.113s | 89K/s | 1.48x slower |

## Key Findings

### Why Simple Formulas Are Faster

**Simple formulas** (`=A{i}+B{i}`, `=A{i}*2`):
- Processed entirely in SQL as bulk column operations
- Single query processes entire column
- DuckDB's vectorized execution engine optimizes these operations
- **No per-row scalar evaluation needed**

**Complex formulas** (`=IF(A{i}="x", B{i}*1.1, B{i}*0.9)`):
- Require Phase 1: SQL aggregate extraction
- Require Phase 2: Per-row scalar evaluation with numexpr
- Each row must be evaluated individually
- **3.5-4.8x slower than simple formulas**

### Why numexpr Matters

The numexpr optimization provides:
1. **1.48x faster** than the old `eval()` approach
2. **Better security** - No arbitrary code execution
3. **Numeric optimization** - Uses vectorized operations under the hood

Without numexpr, complex formulas would be even slower!

## Performance Hierarchy

```
Fastest → Slowest:
┌─────────────────────────────────────────────────────────┐
│ 1. Simple Formulas (Pure SQL)                          │
│    • Entire column in single query                      │
│    • 24-30K rows/s                                      │
├─────────────────────────────────────────────────────────┤
│ 2. Complex Formulas (SQL + numexpr)                     │
│    • Per-row scalar evaluation                          │
│    • 6K rows/s                                          │
├─────────────────────────────────────────────────────────┤
│ 3. Complex Formulas (SQL + eval) [theoretical]         │
│    • Would be ~1.5x slower than numexpr                 │
│    • ~4K rows/s (estimated)                             │
└─────────────────────────────────────────────────────────┘
```

## Implications

1. **For simple formulas**: The pure SQL approach is optimal
2. **For complex formulas**: numexpr is essential for acceptable performance
3. **Benchmark design**: Need separate benchmarks for simple vs complex formulas
4. **Real-world workloads**: Most spreadsheets have mixed formulas, so both code paths matter

## Running the Benchmarks

```bash
# Generate test files
python3 generate_complex_test_files.py

# Run complex formula benchmark
python3 measure_complex_formulas.py test_files_complex/complex_10000.xlsx

# Run comparison
bash run_comparison.sh
```

## Files

- `generate_complex_test_files.py` - Generate complex formula test files
- `measure_complex_formulas.py` - Benchmark complex formulas
- `run_complex_benchmark.sh` - Run all complex formula benchmarks
- `run_comparison.sh` - Compare simple vs complex formulas
- `test_formula_evaluator.py` - FormulaEvaluator with numexpr optimization
