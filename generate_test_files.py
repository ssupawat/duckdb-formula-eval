#!/usr/bin/env python3
"""
Generate test Excel files matching the benchmark format.

Test formats:
- Standard: A=i, B=i*2, C=A{i}+B{i}
- Two-sheet: Sheet1 A=i, Sheet2 A=Sheet1!A{i}, B=A{i}*2
"""

import openpyxl
import sys
from pathlib import Path


def create_standard_test(rows: int, output_path: Path):
    """Create standard single-sheet test (matches benchmark format)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Sheet1'

    # Header
    ws['A1'] = 'A'
    ws['B1'] = 'B'
    ws['C1'] = 'C = A + B'

    # Data with formulas
    for i in range(2, rows + 2):
        ws[f'A{i}'] = i
        ws[f'B{i}'] = i * 2
        ws[f'C{i}'] = f'=A{i}+B{i}'

    wb.save(output_path)
    print(f'Generated {output_path.name}')


def create_two_sheet_test(rows: int, output_path: Path):
    """Create two-sheet test with cross-sheet references (matches benchmark format)."""
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = 'Sheet1'

    # Sheet1: Source data
    ws1['A1'] = 'Value'
    for i in range(2, rows + 2):
        ws1[f'A{i}'] = i

    # Sheet2: References Sheet1
    ws2 = wb.create_sheet('Sheet2')
    ws2['A1'] = 'From Sheet1'
    ws2['B1'] = 'Doubled'

    for i in range(2, rows + 2):
        ws2[f'A{i}'] = f'=Sheet1!A{i}'
        ws2[f'B{i}'] = f'=A{i}*2'

    wb.save(output_path)
    print(f'Generated {output_path.name}')


def main():
    # Standard tests: 10K, 50K, 100K, 200K
    test_dir = Path('test_files')
    test_dir.mkdir(exist_ok=True)

    if len(sys.argv) > 1:
        rows = int(sys.argv[1])
        create_standard_test(rows, test_dir / f'test_{rows}.xlsx')
        create_two_sheet_test(rows, test_dir / f'test_2sheet_{rows}.xlsx')
    else:
        # Generate all benchmark test files
        for n in [10000, 50000, 100000, 200000]:
            create_standard_test(n, test_dir / f'test_{n}.xlsx')

        for n in [10000, 100000, 500000]:
            create_two_sheet_test(n, test_dir / f'test_2sheet_{n}.xlsx')

        # Max rows test (commented out - takes too long)
        # create_standard_test(1048576, test_dir / 'test_max.xlsx')


if __name__ == '__main__':
    main()
