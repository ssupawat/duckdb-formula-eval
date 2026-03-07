#!/bin/bash
echo "=== Complex Formula Benchmark (numexpr Performance) ==="
echo ""

echo "Generating test files..."
python3 generate_complex_test_files.py

echo ""
echo "=== Single Sheet (IF Formulas) ==="
printf "%-10s %-20s %-20s\n" "Rows" "Time (s)" "Peak (MB)"
echo "---------------------------------------------"

for n in 10000 50000 100000; do
    result=$(python3 measure_complex_formulas.py "test_files_complex/complex_${n}.xlsx")
    time=$(echo "$result" | python3 -c "import sys, json; print(json.load(sys.stdin)['timeSeconds'])")
    peak=$(echo "$result" | python3 -c "import sys, json; print(json.load(sys.stdin)['peakMemoryMb'])")
    printf "%-10s %-20s %-20s\n" "$n" "$time" "$peak"
done

echo ""
echo "=== Two Sheets (VLOOKUP Formulas) ==="
printf "%-10s %-20s %-20s\n" "Rows" "Time (s)" "Peak (MB)"
echo "---------------------------------------------"

for n in 10000 50000 100000; do
    result=$(python3 measure_complex_formulas.py "test_files_complex/complex_2sheet_${n}.xlsx")
    time=$(echo "$result" | python3 -c "import sys, json; print(json.load(sys.stdin)['timeSeconds'])")
    peak=$(echo "$result" | python3 -c "import sys, json; print(json.load(sys.stdin)['peakMemoryMb'])")
    printf "%-10s %-20s %-20s\n" "$n" "$time" "$peak"
done
