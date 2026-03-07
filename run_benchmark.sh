#!/bin/bash
# Run DuckDB benchmark and compare with JS/LibreOffice results

set -e

GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[0;33m'
NC='\033[0m'

echo -e "${BLUE}=== DuckDB Formula Evaluation Benchmark (Optimized SQL) ===${NC}\n"

# Generate test files (matching benchmark format)
echo -e "${GREEN}Generating test files...${NC}"
python3 generate_test_files.py

# Standard tests
echo -e "\n${BLUE}=== Standard Tests (1 Sheet) ===${NC}"
echo "Rows     DuckDB Time (s)    DuckDB Peak (MB)"
echo "─────────────────────────────────────────────"

for n in 10000 50000 100000 200000; do
    result=$(python3 measure_duckdb_optimized.py "test_files/test_${n}.xlsx" 2>&1)
    time=$(echo "$result" | python3 -c "import sys, json; print(json.load(sys.stdin)['timeSeconds'])")
    peak=$(echo "$result" | python3 -c "import sys, json; print(json.load(sys.stdin)['peakTotalMB'])")
    printf "%7s       %6.3f             %7.2f\n" "$n" "$time" "$peak"
done

# Two-sheet tests
echo -e "\n${BLUE}=== Two Sheets (Cross-Sheet References) ===${NC}"
echo "Rows     DuckDB Time (s)    DuckDB Peak (MB)"
echo "─────────────────────────────────────────────"

for n in 10000 100000 500000; do
    result=$(python3 measure_duckdb_optimized.py "test_files/test_2sheet_${n}.xlsx" 2>&1)
    time=$(echo "$result" | python3 -c "import sys, json; print(json.load(sys.stdin)['timeSeconds'])")
    peak=$(echo "$result" | python3 -c "import sys, json; print(json.load(sys.stdin)['peakTotalMB'])")
    printf "%7s       %6.3f             %7.2f\n" "$n" "$time" "$peak"
done

echo -e "\n${BLUE}=== Comparison Reference ===${NC}"
echo "JS/LibreOffice results from lo-vs-xlsx-calc-formula-eval-benchmark:"
echo ""
echo "Standard Tests:"
echo "Rows     JS Time (s)    LO Time (s)    JS Peak (MB)    LO Peak (MB)"
echo "───────────────────────────────────────────────────────────────────"
echo " 10K        0.14            1.01            109               222"
echo " 50K        0.46            0.88            158               223"
echo "100K        0.91            1.45            219               283"
echo "200K        1.93            2.04            339               405"
echo ""
echo "Two Sheets:"
echo "Rows     JS Time (s)    LO Time (s)    JS Peak (MB)    LO Peak (MB)"
echo "───────────────────────────────────────────────────────────────────"
echo " 10K        0.12            0.69            108               233"
echo "100K        0.78            1.81            189               376"
echo "500K        4.74            7.04            494             1,186"
