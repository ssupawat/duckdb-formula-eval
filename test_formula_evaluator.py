#!/usr/bin/env python3
"""
Tests for FormulaEvaluator

Run comprehensive tests for Excel formula evaluation including:
- Pure aggregates: SUM, AVERAGE, MAX, MIN, COUNTIF, SUMIF
- Pure scalar: arithmetic, IF statements
- Nested: aggregate inside IF, IF with aggregate conditions
- Arithmetic on aggregates: SUM(D:D)*0.1
- Cross-sheet VLOOKUP
"""

import re
import duckdb
import pandas as pd
import openpyxl
from dataclasses import dataclass
from pathlib import Path
from typing import Dict

from formula_evaluator import FormulaEvaluator


@dataclass
class TestCase:
    formula: str
    output_col: str
    row_ctx: Dict[str, float] = None
    expected: float = None


TEST_CASES = [
    # ── Pure aggregate ──────────────────────────────────────────
    {"formula": "=SUM(D:D)",            "output_col": "total_amount"},
    {"formula": "=AVERAGE(D:D)",        "output_col": "avg_amount"},
    {"formula": "=MAX(D:D)",            "output_col": "max_amount"},
    {"formula": '=SUMIF(C:C,"x",D:D)', "output_col": "sum_x"},
    {"formula": '=COUNTIF(C:C,"x")',   "output_col": "count_x"},

    # ── Pure scalar (ต้องส่ง row_ctx ด้วย) ─────────────────────
    {"formula": "=D1*1.07",                        "output_col": "amount_with_vat",  "row_ctx": {"D1": 100.0}},
    {"formula": "=IF(D1>80, D1*1.1, D1*0.9)",      "output_col": "adjusted_amount",  "row_ctx": {"D1": 100.0}},
    {"formula": "=IF(D1>80, D1*1.1, D1*0.9)",      "output_col": "adjusted_amount",  "row_ctx": {"D1": 50.0}},

    # ── Nested: aggregate inside IF ─────────────────────────────
    {"formula": '=IF(SUMIF(C:C,"x",D:D)>100, D1*1.07, 0)',               "output_col": "conditional_vat",   "row_ctx": {"D1": 100.0}},
    {"formula": '=IF(SUMIF(C:C,"x",D:D)/COUNTIF(C:C,"x")>50, D1*2, D1)', "output_col": "bonus_amount",       "row_ctx": {"D1": 100.0}},

    # ── Arithmetic on aggregate ─────────────────────────────────
    {"formula": "=SUM(D:D)*0.1",        "output_col": "ten_pct_of_total"},
    {"formula": "=AVERAGE(D:D)*1.2",    "output_col": "avg_markup"},

    # ── Cross-sheet VLOOKUP ─────────────────────────────────────
    {"formula": "=VLOOKUP(A1,Sheet2!A:B,2,0)", "output_col": "label", "row_ctx": {"A1": "A"}},
    {"formula": "=VLOOKUP(A1,Sheet2!A:B,2,0)", "output_col": "label", "row_ctx": {"A1": "B"}},
]


def create_test_excel():
    """Create a test Excel file with sample data for testing."""
    wb = openpyxl.Workbook()

    # Sheet1: Main data
    ws1 = wb.active
    ws1.title = 'Sheet1'
    ws1['A1'] = 'Key'
    ws1['B1'] = 'Name'
    ws1['C1'] = 'Category'
    ws1['D1'] = 'Amount'

    # Add test data
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


def main():
    """Run all test cases."""
    print("=" * 70)
    print("Comprehensive Formula Evaluation Test")
    print("=" * 70)

    # Create test Excel file
    print("\n1. Creating test Excel file...")
    test_file = create_test_excel()
    print(f"   Created: {test_file}")

    # Load Excel into DuckDB
    print("\n2. Loading data into DuckDB...")
    conn = duckdb.connect(':memory:')

    excel_file = pd.ExcelFile(test_file, engine='openpyxl')
    sheets_data = {}

    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(excel_file, sheet_name=sheet_name, header=0, engine='openpyxl')
        # Normalize column names
        df.columns = [str(c).lower().replace(' ', '_') for c in df.columns]
        table_name = sheet_name.lower().replace(' ', '_')
        sheets_data[table_name] = df
        conn.register(table_name, df)
        print(f"   Loaded {sheet_name}: {len(df)} rows, columns: {list(df.columns)}")

    # Create evaluator
    evaluator = FormulaEvaluator(conn, sheets_data)

    # Run test cases
    print("\n3. Running test cases...")
    print("-" * 70)

    for i, test_case in enumerate(TEST_CASES, 1):
        formula = test_case["formula"]
        output_col = test_case["output_col"]
        row_ctx = test_case.get("row_ctx", {})

        try:
            result = evaluator.evaluate_formula(formula, 'sheet1', row_ctx)
            row_ctx_str = f", row_ctx={row_ctx}" if row_ctx else ""
            # Format result based on type
            if isinstance(result, str):
                print(f"✓ Test {i:2d}: {formula:40s} → {result:10s}  [{output_col}{row_ctx_str}]")
            else:
                print(f"✓ Test {i:2d}: {formula:40s} → {result:10.2f}  [{output_col}{row_ctx_str}]")
        except Exception as e:
            print(f"✗ Test {i:2d}: {formula:40s} → ERROR: {e}")

    print("-" * 70)
    print(f"\n4. Test file saved as: {test_file}")
    print("   You can verify formulas manually in Excel.")


if __name__ == "__main__":
    main()
