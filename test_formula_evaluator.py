#!/usr/bin/env python3
"""
Tests for FormulaEvaluator using AAA (Arrange-Act-Assert) pattern.

Test coverage:
- Pure aggregates: SUM, AVERAGE, MAX, MIN, COUNTIF, SUMIF
- Pure scalar: arithmetic, IF statements
- Nested: aggregate inside IF, IF with aggregate conditions
- Arithmetic on aggregates: SUM(D:D)*0.1
- Cross-sheet VLOOKUP
"""

import duckdb
import pandas as pd
import openpyxl
from pathlib import Path
from typing import Dict, Any

from formula_evaluator import FormulaEvaluator


# Test data for Sheet1
# Row | Key  | Name   | Category | Amount
# ----|------|--------|----------|--------
# 2   | A    | Item 1 | x        | 100
# 3   | B    | Item 2 | y        | 200
# 4   | C    | Item 3 | x        | 150
# 5   | D    | Item 4 | x        | 75
# 6   | E    | Item 5 | y        | 300

# Test data for Sheet2 (lookup table)
# Row | Key  | Label
# ----|------|--------
# 2   | A    | Label A
# 3   | B    | Label B
# 4   | C    | Label C

TEST_CASES = [
    # ── Pure aggregate ──────────────────────────────────────────
    {
        "name": "SUM(D:D)",
        "formula": "=SUM(D:D)",
        "row_ctx": {},
        "expected": 825.0,  # 100+200+150+75+300
    },
    {
        "name": "AVERAGE(D:D)",
        "formula": "=AVERAGE(D:D)",
        "row_ctx": {},
        "expected": 165.0,  # 825/5
    },
    {
        "name": "MAX(D:D)",
        "formula": "=MAX(D:D)",
        "row_ctx": {},
        "expected": 300.0,  # max(100,200,150,75,300)
    },
    {
        "name": "SUMIF(C:C,\"x\",D:D)",
        "formula": '=SUMIF(C:C,"x",D:D)',
        "row_ctx": {},
        "expected": 325.0,  # 100+150+75 (rows where C='x')
    },
    {
        "name": "COUNTIF(C:C,\"x\")",
        "formula": '=COUNTIF(C:C,"x")',
        "row_ctx": {},
        "expected": 3.0,  # 3 rows where C='x'
    },

    # ── Pure scalar ─────────────────────────────────────────────
    {
        "name": "D1*1.07 (D1=100)",
        "formula": "=D1*1.07",
        "row_ctx": {"D1": 100.0},
        "expected": 107.0,  # 100*1.07
    },
    {
        "name": "IF(D1>80, D1*1.1, D1*0.9) - TRUE branch (D1=100)",
        "formula": "=IF(D1>80, D1*1.1, D1*0.9)",
        "row_ctx": {"D1": 100.0},
        "expected": 110.0,  # 100>80, so 100*1.1
    },
    {
        "name": "IF(D1>80, D1*1.1, D1*0.9) - FALSE branch (D1=50)",
        "formula": "=IF(D1>80, D1*1.1, D1*0.9)",
        "row_ctx": {"D1": 50.0},
        "expected": 45.0,  # 50<80, so 50*0.9
    },

    # ── Nested: aggregate inside IF ─────────────────────────────
    {
        "name": "IF(SUMIF(C:C,\"x\",D:D)>100, D1*1.07, 0)",
        "formula": '=IF(SUMIF(C:C,"x",D:D)>100, D1*1.07, 0)',
        "row_ctx": {"D1": 100.0},
        "expected": 107.0,  # SUMIF=325>100, so 100*1.07
    },
    {
        "name": "IF(SUMIF/COUNTIF>50, D1*2, D1)",
        "formula": '=IF(SUMIF(C:C,"x",D:D)/COUNTIF(C:C,"x")>50, D1*2, D1)',
        "row_ctx": {"D1": 100.0},
        "expected": 200.0,  # 325/3≈108.3>50, so 100*2
    },

    # ── Arithmetic on aggregate ─────────────────────────────────
    {
        "name": "SUM(D:D)*0.1",
        "formula": "=SUM(D:D)*0.1",
        "row_ctx": {},
        "expected": 82.5,  # 825*0.1
    },
    {
        "name": "AVERAGE(D:D)*1.2",
        "formula": "=AVERAGE(D:D)*1.2",
        "row_ctx": {},
        "expected": 198.0,  # 165*1.2
    },

    # ── Cross-sheet VLOOKUP ─────────────────────────────────────
    {
        "name": "VLOOKUP(A1,Sheet2!A:B,2,0) - Key A",
        "formula": "=VLOOKUP(A1,Sheet2!A:B,2,0)",
        "row_ctx": {"A1": "A"},
        "expected": "Label A",
    },
    {
        "name": "VLOOKUP(A1,Sheet2!A:B,2,0) - Key B",
        "formula": "=VLOOKUP(A1,Sheet2!A:B,2,0)",
        "row_ctx": {"A1": "B"},
        "expected": "Label B",
    },
]


def create_test_excel() -> Path:
    """Create a test Excel file with sample data for testing."""
    wb = openpyxl.Workbook()

    # Sheet1: Main data
    ws1 = wb.active
    ws1.title = 'Sheet1'
    ws1['A1'] = 'Key'
    ws1['B1'] = 'Name'
    ws1['C1'] = 'Category'
    ws1['D1'] = 'Amount'

    test_data = [
        ['A', 'Item 1', 'x', 100],
        ['B', 'Item 2', 'y', 200],
        ['C', 'Item 3', 'x', 150],
        ['D', 'Item 4', 'x', 75],
        ['E', 'Item 5', 'y', 300],
    ]

    for i, row in enumerate(test_data, start=2):
        ws1[f'A{i}'] = row[0]
        ws1[f'B{i}'] = row[1]
        ws1[f'C{i}'] = row[2]
        ws1[f'D{i}'] = row[3]

    # Sheet2: Lookup table
    ws2 = wb.create_sheet('Sheet2')
    ws2['A1'] = 'Key'
    ws2['B1'] = 'Label'

    lookup_data = [
        ['A', 'Label A'],
        ['B', 'Label B'],
        ['C', 'Label C'],
    ]

    for i, row in enumerate(lookup_data, start=2):
        ws2[f'A{i}'] = row[0]
        ws2[f'B{i}'] = row[1]

    output_path = Path('test_formulas.xlsx')
    wb.save(output_path)
    return output_path


def setup_evaluator(test_file: Path) -> FormulaEvaluator:
    """
    Arrange: Set up test fixtures (DuckDB connection, data, evaluator).
    """
    conn = duckdb.connect(':memory:')
    excel_file = pd.ExcelFile(test_file, engine='openpyxl')
    sheets_data = {}

    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(excel_file, sheet_name=sheet_name, header=0, engine='openpyxl')
        df.columns = [str(c).lower().replace(' ', '_') for c in df.columns]
        table_name = sheet_name.lower().replace(' ', '_')
        sheets_data[table_name] = df
        conn.register(table_name, df)

    return FormulaEvaluator(conn, sheets_data)


def run_test(test_case: Dict[str, Any], evaluator: FormulaEvaluator) -> bool:
    """
    Act & Assert: Execute test case and verify result matches expected.
    Returns True if test passes, False otherwise.
    """
    name = test_case["name"]
    formula = test_case["formula"]
    row_ctx = test_case["row_ctx"]
    expected = test_case["expected"]

    try:
        # Act: Execute the formula
        result = evaluator.evaluate_formula(formula, 'sheet1', row_ctx)

        # Assert: Verify result matches expected
        if isinstance(expected, str):
            passed = result == expected
            status = "✓" if passed else "✗"
            result_str = f"'{result}'"
            expected_str = f"'{expected}'"
        else:
            passed = abs(result - expected) < 0.01  # Allow small floating point errors
            status = "✓" if passed else "✗"
            result_str = f"{result:.2f}"
            expected_str = f"{expected:.2f}"

        if passed:
            print(f"{status} {name:50s} → {result_str:15s}")
            return True
        else:
            print(f"{status} {name:50s} → Expected: {expected_str}, Got: {result_str}")
            return False

    except Exception as e:
        print(f"✗ {name:50s} → ERROR: {e}")
        return False


def main():
    """Run all test cases using AAA pattern."""
    print("=" * 80)
    print("FormulaEvaluator Test Suite (AAA Pattern)")
    print("=" * 80)

    # Arrange: Set up test fixtures
    print("\n1. ARRANGE - Creating test fixtures...")
    test_file = create_test_excel()
    print(f"   Created test file: {test_file}")

    evaluator = setup_evaluator(test_file)
    print("   Created FormulaEvaluator with DuckDB connection")

    # Act & Assert: Run all test cases
    print("\n2. ACT & ASSERT - Running test cases...")
    print("-" * 80)

    passed = 0
    failed = 0

    for test_case in TEST_CASES:
        if run_test(test_case, evaluator):
            passed += 1
        else:
            failed += 1

    # Summary
    print("-" * 80)
    print(f"\nResults: {passed} passed, {failed} failed out of {len(TEST_CASES)} tests")

    if failed == 0:
        print("✓ All tests passed!")
    else:
        print(f"✗ {failed} test(s) failed")

    print(f"\n3. Test file saved as: {test_file}")
    print("   You can verify formulas manually in Excel.")

    return failed == 0


if __name__ == "__main__":
    import sys
    sys.exit(0 if main() else 1)
