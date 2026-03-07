#!/bin/bash
# Run DuckDB Formula Evaluation Demo

set -e

# Colors for output
GREEN='\033[0;32m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

echo -e "${BLUE}=== DuckDB Formula Evaluation Demo ===${NC}\n"

# Check dependencies
echo "Checking dependencies..."
python3 -c "import duckdb, openpyxl, psutil" 2>/dev/null || {
    echo "Installing dependencies..."
    pip install duckdb openpyxl psutil
}

# Default row count
ROWS=${1:-10000}

echo -e "${GREEN}Generating test files with $ROWS rows...${NC}"
python3 generate_test_files.py "$ROWS"

echo -e "\n${GREEN}Running single-sheet benchmark...${NC}"
python3 measure_duckdb.py "test_files/test_single_${ROWS}.xlsx"

echo -e "\n${GREEN}Running two-sheet benchmark...${NC}"
python3 measure_duckdb.py "test_files/test_twosheet_${ROWS}.xlsx"

echo -e "\n${BLUE}=== Demo Complete ===${NC}"
echo "Output files saved as: output_*.xlsx"
