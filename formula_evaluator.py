#!/usr/bin/env python3
"""
DuckDB Excel Formula Evaluator (POC)

Simple POC for multi-step pipeline integration.
Evaluates Excel formulas using pure DuckDB SQL.

Supported formula types:
- Pure aggregates: SUM, AVERAGE, MAX, MIN, COUNTIF, SUMIF
- Scalar arithmetic: Basic math operations on cell references
- IF statements: Conditional formulas with nested conditions
- Nested formulas: Aggregates inside IF statements, IF with aggregate conditions
- Arithmetic on aggregates: SUM(D:D)*0.1
- Cross-sheet VLOOKUP
- LEN function: LEN(D:D) or LEN({ColumnName})
- Column name references: {ColumnName} syntax
"""

import re
import duckdb
import sys
from typing import Any, Dict, List, Tuple, Optional, Union


class FormulaEvaluator:
    """
    Evaluate Excel formulas using pure DuckDB SQL.

    This evaluator expects data to already exist in DuckDB tables.
    Formula results are written directly to DuckDB tables - no DataFrames needed.

    Supported formula types:
    - Pure aggregates: SUM, AVERAGE, MAX, MIN, COUNTIF, SUMIF
    - Scalar arithmetic: Basic math operations on cell references
    - IF statements: Conditional formulas with nested conditions
    - Nested formulas: Aggregates inside IF statements, IF with aggregate conditions
    - Arithmetic on aggregates: SUM(D:D)*0.1
    - Cross-sheet VLOOKUP
    """

    def __init__(self, conn: duckdb.DuckDBPyConnection):
        """
        Initialize the evaluator.

        Args:
            conn: DuckDB connection with tables already registered
                  (Data loading is handled separately, not by the evaluator)
        """
        self.conn = conn
        self.last_sql = None  # Store last generated SQL for debugging

        # Cache column names from DuckDB information_schema
        # Structure: {table_name: [col_name1, col_name2, ...]}
        self._column_cache: Dict[str, List[str]] = {}

        # Formula metadata storage for recalculation
        # Structure: {table_name: {target_column: formula}}
        self.formulas: Dict[str, Dict[str, str]] = {}

    # ========================================================================
    # PURE SQL CONVERSION METHODS
    # ========================================================================

    def excel_to_sql(self, formula: str, sheet_name: str, row_ctx: Dict[str, float] = None) -> str:
        """
        Convert Excel formula to pure DuckDB SQL.

        Conversion pipeline order (critical):
        1. Cross-sheet cell references → SQL (before string literal conversion)
        2. VLOOKUP → SQL subqueries (before string literal conversion!)
        3. String literals (double quotes) → SQL (single quotes)
        4. Aggregates → SQL subqueries
        5. LEN function → SQL LENGTH (handles {ColumnName} internally)
        6. {ColumnName} → SQL column references (for non-LEN contexts)
        7. IF → CASE expressions
        8. Cell references → scalar values
        9. Operators → SQL operators

        Args:
            formula: Excel formula (with or without leading =)
            sheet_name: Name of the sheet containing the formula
            row_ctx: Optional row context for cell references

        Returns:
            SQL SELECT statement that evaluates the formula
        """
        # Remove leading = and whitespace
        expr = formula.lstrip('=').strip()
        table_name = sheet_name.lower().replace(' ', '_')

        # Step 1: Cross-sheet cell references → SQL (before string literal conversion)
        # Use backticks for column names to avoid conflicts with Excel string literals
        expr = self._convert_cross_sheet_ref_to_sql(expr, sheet_name, row_ctx)

        # Step 2: VLOOKUP → SQL subqueries (before string literal conversion!)
        expr = self._convert_vlookup_to_sql(expr, sheet_name)

        # Step 3: Convert string literals (Excel "text" → SQL 'text')
        expr = self._convert_string_literals(expr)

        # Step 4: Aggregates → SQL subqueries
        expr = self._convert_aggregates_to_sql(expr, sheet_name)

        # Step 5: LEN function → SQL LENGTH (handles {ColumnName} internally)
        expr = self._convert_len_to_sql(expr, table_name, row_ctx)

        # Step 6: {ColumnName} → SQL column references (for non-LEN contexts)
        expr = self._convert_braced_column_to_sql(expr, table_name)

        # Step 7: IF → CASE expressions
        expr = self._convert_if_to_sql(expr)

        # Step 8: Cell references → scalar values
        if row_ctx:
            expr = self._substitute_cell_references(expr, row_ctx)

        # Step 9: Operators → SQL operators
        expr = self._convert_operators(expr)

        return f"SELECT {expr}"

    def _convert_string_literals(self, formula: str) -> str:
        """Convert Excel string literals (double quotes) to SQL (single quotes)."""
        # Excel: "text" → SQL: 'text'
        return formula.replace('"', "'")

    def _convert_if_to_sql(self, formula: str) -> str:
        """Convert Excel IF statements to SQL CASE expressions."""
        # Pattern: IF(condition, true_value, false_value)
        result = []
        i = 0
        expr = formula

        while i < len(expr):
            if expr[i:i+3] == 'IF(' and (i == 0 or not expr[i-1].isalnum()):
                # Find matching closing parenthesis
                depth = 1
                j = i + 3
                while j < len(expr) and depth > 0:
                    if expr[j] == '(':
                        depth += 1
                    elif expr[j] == ')':
                        depth -= 1
                    j += 1

                if depth == 0:
                    # Extract IF content and parse parameters
                    if_content = expr[i+3:j-1]
                    params = self._split_if_params(if_content)

                    if len(params) == 3:
                        # Check if branches have mixed types (string and numeric)
                        has_string_literal = any(
                            (p.strip().startswith("'") and p.strip().endswith("'")) or
                            (p.strip().startswith('"') and p.strip().endswith('"'))
                            for p in [params[1], params[2]]
                        )

                        if has_string_literal:
                            # Wrap both branches in CAST to VARCHAR for type compatibility
                            case_expr = f"CASE WHEN {params[0]} THEN CAST({params[1]} AS VARCHAR) ELSE CAST({params[2]} AS VARCHAR) END"
                        else:
                            case_expr = f"CASE WHEN {params[0]} THEN {params[1]} ELSE {params[2]} END"
                        result.append(case_expr)
                        i = j
                        continue

            result.append(expr[i])
            i += 1

        return ''.join(result)

    def _split_if_params(self, s: str) -> list:
        """Split IF parameters, respecting nested parentheses and strings."""
        params = []
        current = []
        depth = 0
        in_string = False
        string_char = None

        for char in s:
            if char in ('"', "'") and (in_string is False or string_char == char):
                in_string = not in_string
                if in_string:
                    string_char = char
                else:
                    string_char = None
                current.append(char)
            elif in_string:
                current.append(char)
            elif char == '(':
                depth += 1
                current.append(char)
            elif char == ')':
                depth -= 1
                current.append(char)
            elif char == ',' and depth == 0:
                params.append(''.join(current).strip())
                current = []
            else:
                current.append(char)

        if current:
            params.append(''.join(current).strip())

        return params

    def _convert_aggregates_to_sql(self, formula: str, sheet_name: str) -> str:
        """Convert Excel aggregate functions to SQL subqueries."""
        table_name = sheet_name.lower().replace(' ', '_')
        # Check if table exists in DuckDB
        try:
            self.conn.execute(f'SELECT 1 FROM {table_name} LIMIT 1')
        except Exception:
            return formula  # Table doesn't exist, return formula unchanged

        # Handle COUNT(D:D) pattern
        formula = re.sub(
            r'COUNT\(([A-Z]):([A-Z])\)',
            lambda m: self._count_to_sql(m, table_name),
            formula
        )

        # Handle SUM(D:D) pattern
        formula = re.sub(
            r'SUM\(([A-Z]):([A-Z])\)',
            lambda m: self._sum_to_sql(m, table_name),
            formula
        )

        # Handle AVERAGE(D:D) pattern
        formula = re.sub(
            r'AVERAGE\(([A-Z]):([A-Z])\)',
            lambda m: self._average_to_sql(m, table_name),
            formula
        )

        # Handle MAX(D:D) pattern
        formula = re.sub(
            r'MAX\(([A-Z]):([A-Z])\)',
            lambda m: self._max_to_sql(m, table_name),
            formula
        )

        # Handle MIN(D:D) pattern
        formula = re.sub(
            r'MIN\(([A-Z]):([A-Z])\)',
            lambda m: self._min_to_sql(m, table_name),
            formula
        )

        # Handle COUNTIF(C:C,">100") pattern - with comparison operators (FIRST!)
        formula = re.sub(
            r'COUNTIF\(([A-Z]):([A-Z]),\s*"((?:[<>]=?|<>|=)(?:\d+(?:\.\d+)?))"\)',
            lambda m: self._countif_op_to_sql(m, table_name),
            formula
        )

        # Handle COUNTIF(C:C,'>100') pattern - single quotes with operators
        formula = re.sub(
            r"COUNTIF\(([A-Z]):([A-Z]),\s*'((?:[<>]=?|<>|=)(?:\d+(?:\.\d+)?))'\)",
            lambda m: self._countif_op_to_sql(m, table_name),
            formula
        )

        # Handle COUNTIF(C:C,"") pattern - empty criteria (BEFORE simple equality!)
        formula = re.sub(
            r'COUNTIF\(([A-Z]):([A-Z]),\s*""\)',
            lambda m: self._countif_empty_to_sql(m, table_name),
            formula
        )

        # Handle COUNTIF(C:C,'') pattern - empty criteria with single quotes
        formula = re.sub(
            r"COUNTIF\(([A-Z]):([A-Z]),\s*''\)",
            lambda m: self._countif_empty_to_sql(m, table_name),
            formula
        )

        # Handle COUNTIF(C:C,"x") pattern - simple equality
        formula = re.sub(
            r'COUNTIF\(([A-Z]):([A-Z]),\s*"([^"]*)"\)',
            lambda m: self._countif_to_sql(m, table_name),
            formula
        )

        # Handle COUNTIF(C:C,'x') pattern (single quotes)
        formula = re.sub(
            r"COUNTIF\(([A-Z]):([A-Z]),\s*'([^']*)'\)",
            lambda m: self._countif_to_sql(m, table_name),
            formula
        )

        # Handle SUMIF(C:C,">100",D:D) pattern - with comparison operators (FIRST!)
        formula = re.sub(
            r'SUMIF\(([A-Z]):([A-Z]),\s*"((?:[<>]=?|<>|=)(?:\d+(?:\.\d+)?))",\s*([A-Z]):([A-Z])\)',
            lambda m: self._sumif_op_to_sql(m, table_name),
            formula
        )

        # Handle SUMIF(C:C,'>100',D:D) pattern - single quotes with operators
        formula = re.sub(
            r"SUMIF\(([A-Z]):([A-Z]),\s*'((?:[<>]=?|<>|=)(?:\d+(?:\.\d+)?))',\s*([A-Z]):([A-Z])\)",
            lambda m: self._sumif_op_to_sql(m, table_name),
            formula
        )

        # Handle SUMIF(C:C,"",D:D) pattern - empty criteria (BEFORE simple equality!)
        formula = re.sub(
            r'SUMIF\(([A-Z]):([A-Z]),\s*"",\s*([A-Z]):([A-Z])\)',
            lambda m: self._sumif_empty_to_sql(m, table_name),
            formula
        )

        # Handle SUMIF(C:C,'',D:D) pattern - empty criteria with single quotes
        formula = re.sub(
            r"SUMIF\(([A-Z]):([A-Z]),\s*'',\s*([A-Z]):([A-Z])\)",
            lambda m: self._sumif_empty_to_sql(m, table_name),
            formula
        )

        # Handle SUMIF(C:C,"x",D:D) pattern - simple equality
        formula = re.sub(
            r'SUMIF\(([A-Z]):([A-Z]),\s*"([^"]*)",\s*([A-Z]):([A-Z])\)',
            lambda m: self._sumif_to_sql(m, table_name),
            formula
        )

        # Handle SUMIF(C:C,'x',D:D) pattern (single quotes)
        formula = re.sub(
            r"SUMIF\(([A-Z]):([A-Z]),\s*'([^']*)',\s*([A-Z]):([A-Z])\)",
            lambda m: self._sumif_to_sql(m, table_name),
            formula
        )

        return formula

    def _convert_cross_sheet_ref_to_sql(self, formula: str, sheet_name: str, row_ctx: Dict[str, float] = None) -> str:
        """
        Convert cross-sheet cell references to SQL.

        Pattern: SheetName!A2, Sheet2!B5, etc.

        For single-cell evaluation (excel_to_sql):
        - If row_ctx contains the cell value, use it
        - Otherwise, query the database for the specific row

        For column evaluation (apply_formula_to_column):
        - Handled by _build_vectorized_sql_expression with row-by-row logic
        """
        # Pattern: SheetName!CellRef (e.g., Sheet2!A2, Data!B5)
        pattern = r'([A-Za-z0-9_]+)!([A-Z])\d+'

        def replace_ref(m):
            target_sheet = m.group(1)
            col_letter = m.group(2)
            cell_ref = m.group(0)  # Full reference like Sheet2!A2

            target_table = target_sheet.lower().replace(' ', '_')

            # Check if target table exists
            try:
                self.conn.execute(f'SELECT 1 FROM {target_table} LIMIT 1')
            except Exception:
                # Table doesn't exist - for excel_to_sql, we need to handle this
                # Raise an error that will be caught by the test
                raise duckdb.InvalidInputException(f"Table '{target_table}' does not exist for cross-sheet reference '{cell_ref}'")

            # For single-cell evaluation with row_ctx, check if the value is provided
            if row_ctx and cell_ref in row_ctx:
                return str(row_ctx[cell_ref])

            # For apply_formula_to_column (column operations), this is handled
            # by _build_vectorized_sql_expression. For excel_to_sql without row_ctx,
            # we need a different approach.
            # Since we're generating SQL for a single cell, we can't do row-by-row here.
            # Return a placeholder that indicates this needs row context.
            # For now, let's query the first row's value (LIMIT 1)
            col_name = self._get_column_name(col_letter, target_table)
            if col_name:
                # Use no quotes - DuckDB accepts unquoted lowercase identifiers
                return f'(SELECT {col_name} FROM {target_table} LIMIT 1)'
            return cell_ref  # Return original if column not found

        return re.sub(pattern, replace_ref, formula)

    def _convert_vlookup_to_sql(self, formula: str, sheet_name: str) -> str:
        """Convert VLOOKUP to SQL subquery."""
        # VLOOKUP("value", Sheet2!A:B, 2, 0) or VLOOKUP(A1, Sheet2!A:B, 2, 0)
        pattern = r'VLOOKUP\(("([^"]*)"|([^,]+)),\s*([A-Za-z0-9_]+)!([A-Z]):([A-Z]),\s*(\d+),\s*([01])\)'

        def replace_vlookup(m):
            lookup_value = m.group(2) or m.group(3)  # Either "value" or A1
            target_sheet = m.group(4)
            col_start = m.group(5)
            col_end = m.group(6)
            col_index = int(m.group(7))
            range_lookup = int(m.group(8))

            target_table = target_sheet.lower().replace(' ', '_')

            # Check if target table exists
            try:
                self.conn.execute(f'SELECT 1 FROM {target_table} LIMIT 1')
            except Exception:
                return '0'

            # Ensure index exists on lookup column (optimization for VLOOKUP)
            lookup_col = self._get_column_name(col_start, target_table)
            if lookup_col:
                self._ensure_index(target_table, lookup_col)

            # Get column names
            lookup_col = self._get_column_name(col_start, target_table)
            return_col_name = chr(ord(col_start) + col_index - 1)
            return_col = self._get_column_name(return_col_name, target_table)

            if not lookup_col or not return_col:
                return '0'

            # If lookup_value is a cell reference, keep it for substitution
            if re.match(r'^[A-Z]\d+$', lookup_value):
                lookup_sql = lookup_value
            elif re.match(r'^\d+(?:\.\d+)?$', lookup_value):
                # Numeric literal
                lookup_sql = lookup_value
            else:
                # String literal - Excel uses double quotes, SQL uses single quotes
                lookup_sql = f"'{lookup_value}'"

            if range_lookup == 0:
                # Exact match
                sql = f"(SELECT COALESCE((SELECT {return_col} FROM {target_table} WHERE {lookup_col} = {lookup_sql} LIMIT 1), NULL))"
            else:
                # Approximate match (range_lookup=1): Find largest value ≤ lookup_value
                sql = f"(SELECT COALESCE((SELECT {return_col} FROM {target_table} WHERE {lookup_col} <= {lookup_sql} ORDER BY {lookup_col} DESC LIMIT 1), NULL))"
            return sql

        return re.sub(pattern, replace_vlookup, formula)

    def _ensure_index(self, table_name: str, column_name: str) -> None:
        """
        Ensure an index exists on a column for faster VLOOKUP.

        This is a simple optimization that turns O(n) lookups into O(log n).
        """
        index_name = f'idx_{table_name}_{column_name}'
        try:
            # Check if index exists
            self.conn.execute(f"SELECT * FROM information_schema.indexes WHERE index_name = '{index_name}'")
        except Exception:
            # Create index if it doesn't exist
            try:
                self.conn.execute(f'CREATE INDEX {index_name} ON {table_name}("{column_name}")')
            except Exception as e:
                print(f"Warning: Failed to create index '{index_name}' on '{table_name}': {e}", file=sys.stderr)

    def _substitute_cell_references(self, formula: str, row_ctx: Dict[str, float]) -> str:
        """Substitute cell references with scalar values from row context."""
        result = []
        i = 0

        while i < len(formula):
            if formula[i].isalpha() and i + 1 < len(formula) and formula[i + 1].isdigit():
                # Found cell reference like A1, B2
                cell_ref = formula[i:i + 2]

                if cell_ref in row_ctx:
                    value = row_ctx[cell_ref]
                    if isinstance(value, str):
                        result.append(f"'{value}'")
                    else:
                        result.append(str(value))
                    i += 2
                    continue

            result.append(formula[i])
            i += 1

        return ''.join(result)

    def _convert_operators(self, formula: str) -> str:
        """Convert Excel operators to SQL operators."""
        # Excel <> → SQL !=
        formula = formula.replace('<>', '!=')
        # Excel = for comparison → SQL == (but need to be careful not to replace in CASE)
        # This is handled in the context of SQL expressions
        return formula

    def _convert_braced_column_to_sql(self, formula: str, table_name: str) -> str:
        """
        Convert {ColumnName} syntax to SQL column reference.

        Pattern: {TaxID} → "taxid" (actual column name in DuckDB)
        Uses case-insensitive header lookup to find the column.

        Note: For LEN({ColumnName}), the LEN function handles braces internally.
        This converts braces in other contexts (future use).
        """
        pattern = r'\{([A-Za-z_][A-Za-z0-9_]*)\}'

        def replace_braced(m):
            header_name = m.group(1)
            col_name = self._get_column_by_header(header_name, table_name)
            if col_name:
                return f'"{col_name}"'
            return m.group(0)

        return re.sub(pattern, replace_braced, formula)

    def _convert_len_to_sql(self, formula: str, table_name: str, row_ctx: Dict[str, float] = None) -> str:
        """
        Convert LEN function to SQL LENGTH function.

        Patterns:
        - LEN(D:D) → LENGTH("value") when row_ctx provided, LENGTH("column3") for aggregate
        - LEN({TaxID}) → LENGTH("value") when row_ctx provided, LENGTH("taxid") for aggregate
        - LEN(A2) → LENGTH("value") when row_ctx provided, LENGTH("column0") otherwise
        """
        # Helper to find cell value in row_ctx for a given column letter
        def get_cell_value_for_column(col_letter: str) -> Any:
            if not row_ctx:
                return None
            # Find the first cell in row_ctx that starts with the column letter
            for cell_ref, value in row_ctx.items():
                if cell_ref.startswith(col_letter) and cell_ref[len(col_letter):].isdigit():
                    return value
            return None

        # Helper to find cell value in row_ctx for a given column name
        def get_cell_value_for_column_by_name(col_name: str) -> Any:
            if not row_ctx:
                return None
            # Get column letter from column name
            col_letter = self._get_column_letter_for_name(col_name, table_name)
            if col_letter:
                return get_cell_value_for_column(col_letter)
            return None

        # Pattern: LEN(column_ref) where column_ref can be D:D, {Name}, or A2
        def replace_len(m):
            col_ref = m.group(1)

            # Check if it's a braced column reference {TaxID}
            # This must be checked BEFORE replacing braces in general formula
            braced_match = re.match(r'^\{([A-Za-z_][A-Za-z0-9_]*)\}$', col_ref)
            if braced_match:
                header_name = braced_match.group(1)
                col_name = self._get_column_by_header(header_name, table_name)
                if col_name:
                    # Check if we have a row_ctx value for this column
                    cell_value = get_cell_value_for_column_by_name(col_name)
                    if cell_value is not None:
                        if isinstance(cell_value, str):
                            return f"LENGTH('{cell_value}')"
                        return f"LENGTH('{cell_value}')"
                    return f'LENGTH("{col_name}")'
                return f'LENGTH(NULL)'  # Column not found

            # Check if it's a column range D:D
            range_match = re.match(r'^([A-Z]):([A-Z])$', col_ref)
            if range_match:
                col_letter = range_match.group(1)
                # Check if we have a row_ctx value for this column
                cell_value = get_cell_value_for_column(col_letter)
                if cell_value is not None:
                    if isinstance(cell_value, str):
                        return f"LENGTH('{cell_value}')"
                    return f"LENGTH('{cell_value}')"
                # Fall back to column reference
                col_name = self._get_column_name(col_letter, table_name)
                if col_name:
                    return f'LENGTH("{col_name}")'
                return 'LENGTH(NULL)'  # Column not found

            # Check if it's a cell reference A2
            cell_match = re.match(r'^([A-Z])\d+$', col_ref)
            if cell_match:
                cell_ref = col_ref
                # Check if we have this cell in row_ctx
                if row_ctx and cell_ref in row_ctx:
                    value = row_ctx[cell_ref]
                    if isinstance(value, str):
                        return f"LENGTH('{value}')"
                    return f"LENGTH('{value}')"
                # Fall back to column reference
                col_letter = cell_match.group(1)
                col_name = self._get_column_name(col_letter, table_name)
                if col_name:
                    return f'LENGTH("{col_name}")'
                return 'LENGTH(NULL)'  # Column not found

            # Unknown format, return as-is
            return f'LENGTH({col_ref})'

        return re.sub(r'LEN\(([^)]+)\)', replace_len, formula)

    def _get_column_letter_for_name(self, col_name: str, table_name: str) -> Optional[str]:
        """Get the column letter (A, B, C, etc.) for a given column name."""
        cols = self._get_cached_columns(table_name)
        if col_name in cols:
            return chr(ord('A') + cols.index(col_name))
        return None

    # Aggregate conversion methods
    def _sum_to_sql(self, m: re.Match, table_name: str) -> str:
        col = self._get_column_name(m.group(1), table_name)
        if not col:
            return '0'
        return f"(SELECT COALESCE(SUM(\"{col}\"), 0) FROM {table_name})"

    def _average_to_sql(self, m: re.Match, table_name: str) -> str:
        col = self._get_column_name(m.group(1), table_name)
        if not col:
            return '0'
        return f"(SELECT COALESCE(AVG(\"{col}\"), 0) FROM {table_name})"

    def _max_to_sql(self, m: re.Match, table_name: str) -> str:
        col = self._get_column_name(m.group(1), table_name)
        if not col:
            return '0'
        return f"(SELECT COALESCE(MAX(\"{col}\"), 0) FROM {table_name})"

    def _min_to_sql(self, m: re.Match, table_name: str) -> str:
        col = self._get_column_name(m.group(1), table_name)
        if not col:
            return '0'
        return f"(SELECT COALESCE(MIN(\"{col}\"), 0) FROM {table_name})"

    def _count_to_sql(self, m: re.Match, table_name: str) -> str:
        """Convert COUNT(D:D) to SQL, handling non-existent columns."""
        col = self._get_column_name(m.group(1), table_name)
        if not col:
            return '0'  # Column doesn't exist, return 0
        return f"(SELECT COUNT(*) FROM {table_name})"

    def _countif_to_sql(self, m: re.Match, table_name: str) -> str:
        criteria = m.group(3)
        col = self._get_column_name(m.group(1), table_name)
        if not col:
            return '0'
        return f"(SELECT COUNT(*) FROM {table_name} WHERE \"{col}\" = '{criteria}')"

    def _countif_op_to_sql(self, m: re.Match, table_name: str) -> str:
        """Handle COUNTIF with comparison operators like >100, <=50"""
        criteria = m.group(3)  # e.g., ">100"
        col = self._get_column_name(m.group(1), table_name)
        if not col:
            return '0'
        return f"(SELECT COUNT(*) FROM {table_name} WHERE \"{col}\" {criteria})"

    def _sumif_to_sql(self, m: re.Match, table_name: str) -> str:
        criteria = m.group(3)
        filter_col = self._get_column_name(m.group(1), table_name)
        sum_col = self._get_column_name(m.group(4), table_name)
        if not filter_col or not sum_col:
            return '0'
        return f"(SELECT COALESCE(SUM(\"{sum_col}\"), 0) FROM {table_name} WHERE \"{filter_col}\" = '{criteria}')"

    def _sumif_op_to_sql(self, m: re.Match, table_name: str) -> str:
        """Handle SUMIF with comparison operators like ">100" """
        criteria = m.group(3)  # e.g., ">100"
        filter_col = self._get_column_name(m.group(1), table_name)
        sum_col = self._get_column_name(m.group(5), table_name)
        if not filter_col or not sum_col:
            return '0'
        return f"(SELECT COALESCE(SUM(\"{sum_col}\"), 0) FROM {table_name} WHERE \"{filter_col}\" {criteria})"

    def _countif_empty_to_sql(self, m: re.Match, table_name: str) -> str:
        """Handle COUNTIF with empty string criteria - matches NULL (blank cells)."""
        filter_col = self._get_column_name(m.group(1), table_name)
        if not filter_col:
            return '0'
        return f"(SELECT COUNT(*) FROM {table_name} WHERE \"{filter_col}\" IS NULL)"

    def _sumif_empty_to_sql(self, m: re.Match, table_name: str) -> str:
        """Handle SUMIF with empty string criteria - matches NULL (blank cells)."""
        filter_col = self._get_column_name(m.group(1), table_name)
        sum_col = self._get_column_name(m.group(3), table_name)
        if not filter_col or not sum_col:
            return '0'
        return f"(SELECT COALESCE(SUM(\"{sum_col}\"), 0) FROM {table_name} WHERE \"{filter_col}\" IS NULL)"

    # ========================================================================
    # PATTERN DETECTION & VECTORIZED EVALUATION
    # ========================================================================

    def _parse_formula_pattern(self, formula: str) -> Dict[str, Any]:
        """Detect simple formula patterns for vectorized evaluation."""
        formula_clean = formula.lstrip('=').strip().upper()

        # Pattern: A2+B2 (two columns with operator)
        match = re.match(r'^([A-Z])\d+\s*([+\-*/])\s*([A-Z])\d+$', formula_clean)
        if match:
            return {'type': 'simple', 'col1': match.group(1), 'op': match.group(2), 'col2': match.group(3)}

        # Pattern: A2*2 (column with scalar)
        match = re.match(r'^([A-Z])\d+\s*([+\-*/])\s*(\d+(?:\.\d+)?)$', formula_clean)
        if match:
            return {'type': 'scalar', 'col': match.group(1), 'op': match.group(2), 'value': match.group(3)}

        # Pattern: Sheet1!A2 (cross-sheet reference)
        match = re.match(r'^([A-Za-z0-9_]+)!([A-Z])\d+$', formula_clean)
        if match:
            return {'type': 'cross_sheet', 'sheet': match.group(1), 'col': match.group(2)}

        # === NEW: Pattern for IF statements with column references ===
        # IF(D2>100, D2*1.1, D2) or IF(A2>B2, A2, B2)
        # Pattern: IF(COL1 OP VALUE, COL2 OP2 VALUE2, COL3)
        # where OP is comparison operator, OP2 is arithmetic operator
        match = re.match(
            r'^IF\(([A-Z])\d+([><=!]+)([\d.]+),\s*([A-Z])\d+([+\-*/])([\d.]+),\s*([A-Z])\d+\)$',
            formula_clean
        )
        if match:
            return {
                'type': 'if',
                'col1': match.group(1),           # Condition column
                'op': match.group(2),             # Comparison operator
                'val': match.group(3),            # Comparison value
                'result_col': match.group(4),     # True result column
                'result_op': match.group(5),      # True result operator
                'result_val': match.group(6),     # True result value
                'else_col': match.group(7)        # False result column
            }

        # Also handle IF with column-column comparison: IF(A2>B2, A2, B2)
        match = re.match(
            r'^IF\(([A-Z])\d+([><=!]+)([A-Z])\d+,\s*([A-Z])\d+,\s*([A-Z])\d+\)$',
            formula_clean
        )
        if match:
            return {
                'type': 'if',
                'col1': match.group(1),           # Condition column 1
                'op': match.group(2),             # Comparison operator
                'col2': match.group(3),           # Condition column 2
                'result_col': match.group(4),     # True result column
                'else_col': match.group(5)        # False result column
            }

        # Pattern: LEN({ColumnName})=13 or LEN(D:D)=13
        # Support both {ColumnName} and D:D syntax
        # Extract from original formula (before upper) to preserve case for {ColumnName}
        original_formula = formula.lstrip('=').strip()
        match = re.match('^LEN\\(\\{([A-Za-z_][A-Za-z0-9_]*)\\}\\)\\s*([><=!]+)\\s*(\\d+(?:\\.\\d+)?)$', original_formula, re.IGNORECASE)
        if match:
            return {
                'type': 'len_comparison_braced',
                'header': match.group(1),         # Column header name (preserved case)
                'op': match.group(2),             # Comparison operator
                'value': match.group(3)           # Comparison value
            }

        match = re.match(r'^LEN\(([A-Z]):([A-Z])\)\s*([><=!]+)\s*(\d+(?:\.\d+)?)$', formula_clean)
        if match:
            return {
                'type': 'len_comparison_range',
                'col': match.group(1),            # Column letter
                'op': match.group(3),             # Comparison operator
                'value': match.group(4)           # Comparison value
            }

        return {'type': 'complex'}

    def _get_cached_columns(self, table_name: str) -> List[str]:
        """
        Get cached column names for a table, populating cache if necessary.

        Args:
            table_name: DuckDB table name

        Returns:
            List of column names in order
        """
        if table_name not in self._column_cache:
            try:
                result = self.conn.execute(f"""
                    SELECT column_name
                    FROM information_schema.columns
                    WHERE table_name = '{table_name.lower()}'
                    ORDER BY ordinal_position
                """).fetchall()
                self._column_cache[table_name] = [row[0] for row in result]
            except Exception:
                self._column_cache[table_name] = []
        return self._column_cache[table_name]

    def _get_column_name(self, col_letter: str, table_name: str) -> Optional[str]:
        """
        Map Excel column letter to actual column name using DuckDB information_schema.

        Args:
            col_letter: Excel column letter (e.g., 'A', 'B', 'C')
            table_name: DuckDB table name

        Returns:
            Actual column name or None if not found
        """
        # Check cache first
        if table_name not in self._column_cache:
            try:
                # Query DuckDB's information_schema for column names
                result = self.conn.execute(f"""
                    SELECT column_name
                    FROM information_schema.columns
                    WHERE table_name = '{table_name.lower()}'
                    ORDER BY ordinal_position
                """).fetchall()
                self._column_cache[table_name] = [row[0] for row in result]
            except Exception:
                # Table doesn't exist or error
                self._column_cache[table_name] = []

        columns = self._column_cache[table_name]
        col_idx = ord(col_letter.upper()) - ord('A')

        if 0 <= col_idx < len(columns):
            return columns[col_idx]
        return None

    def _get_column_by_header(self, header_name: str, table_name: str) -> Optional[str]:
        """
        Find column by header name (case-insensitive reverse lookup).

        Tries multiple matching strategies:
        1. Exact match (case-insensitive)
        2. Underscore conversion: "ShortID" → "short_id"

        Args:
            header_name: Column header name to search for (e.g., 'TaxID', 'taxid')
            table_name: DuckDB table name

        Returns:
            Actual DuckDB column name or None if not found
        """
        # Ensure cache is populated
        if table_name not in self._column_cache:
            try:
                result = self.conn.execute(f"""
                    SELECT column_name
                    FROM information_schema.columns
                    WHERE table_name = '{table_name.lower()}'
                    ORDER BY ordinal_position
                """).fetchall()
                self._column_cache[table_name] = [row[0] for row in result]
            except Exception:
                self._column_cache[table_name] = []

        header_lower = header_name.lower()

        # Strategy 1: Direct case-insensitive match
        for col_name in self._column_cache[table_name]:
            if col_name.lower() == header_lower:
                return col_name

        # Strategy 2: Try underscore conversion (PascalCase → snake_case)
        # "ShortID" → "short_id"
        import re
        # Insert underscore before uppercase letters that are followed by lowercase letters
        # or before uppercase letters that follow lowercase letters
        snake_case = re.sub(r'([a-z])([A-Z])', r'\1_\2', header_name).lower()
        # Also handle consecutive uppercase: "TaxID" → "tax_id"
        snake_case = re.sub(r'([A-Z]+)([A-Z][a-z])', r'\1_\2', snake_case).lower()
        for col_name in self._column_cache[table_name]:
            if col_name.lower() == snake_case:
                return col_name

        return None

    # ========================================================================
    # PERSISTENCE METHODS (POC - simplified)
    # ========================================================================
    # PUBLIC API: APPLY FORMULA TO COLUMN (SQL EXECUTION ONLY)
    # ========================================================================

    def apply_formula_to_column(
        self,
        formula: str,
        sheet_name: str,
        target_column: str
    ) -> None:
        """
        Apply formula to column by executing SQL in DuckDB.

        This converts Excel formula to SQL and executes UPDATE directly.
        No evaluation happens in Python - everything runs in DuckDB.

        Args:
            formula: Excel formula (e.g., "=A2+B2", "=IF(A2>100, A2*1.1, A2)")
            sheet_name: Name of the sheet
            target_column: Name of column to store results (must exist)

        Example:
            evaluator = FormulaEvaluator(conn)
            evaluator.apply_formula_to_column('=A2+B2', 'sheet1', 'c')
            # Executes: UPDATE sheet1 SET "c" = "a" + "b"
            # Results stay in DuckDB - no return value
        """
        table_name = sheet_name.lower().replace(' ', '_')

        # Build SQL expression from formula
        pattern = self._parse_formula_pattern(formula)

        # Special handling for cross-sheet references - use row-by-row copy
        if pattern['type'] == 'cross_sheet':
            source_table = pattern['sheet'].lower().replace(' ', '_')
            source_col = self._get_column_name(pattern['col'], source_table)

            if source_col:
                # Get the list of columns in target table
                target_cols_result = self.conn.execute(f"""
                    SELECT column_name
                    FROM information_schema.columns
                    WHERE table_name = '{table_name}'
                    ORDER BY ordinal_position
                """).fetchall()
                target_cols = [row[0] for row in target_cols_result]

                # Build SELECT statement that joins both tables on row position
                # We use row_number() to create a temporary row ID for joining
                select_cols = []
                for col in target_cols:
                    if col == target_column:
                        select_cols.append(f's."{source_col}" AS "{col}"')
                    else:
                        select_cols.append(f't."{col}"')

                # Recreate table with row-by-row joined data
                self.conn.execute(f"""
                    CREATE OR REPLACE TABLE {table_name} AS
                    SELECT {', '.join(select_cols)}
                    FROM (
                        SELECT *, row_number() OVER () AS _rn
                        FROM {source_table}
                    ) s
                    JOIN (
                        SELECT *, row_number() OVER () AS _rn
                        FROM {table_name}
                    ) t ON s._rn = t._rn
                """)
            else:
                raise ValueError(f"Column {pattern['col']} not found in table {source_table}")

            # Store formula metadata for recalculation
            if table_name not in self.formulas:
                self.formulas[table_name] = {}
            self.formulas[table_name][target_column] = formula
            return

        # Regular formula handling
        sql_expr = self._build_vectorized_sql_expression(formula, table_name, pattern)

        # Execute UPDATE directly in DuckDB
        sql = f'UPDATE {table_name} SET "{target_column}" = {sql_expr}'
        self.conn.execute(sql)

        # Store formula metadata for recalculation
        if table_name not in self.formulas:
            self.formulas[table_name] = {}
        self.formulas[table_name][target_column] = formula

    def _build_vectorized_sql_expression(self, formula: str, table_name: str, pattern: Dict[str, Any]) -> str:
        """Build SQL expression from Excel formula pattern."""
        if pattern['type'] == 'simple':
            col1 = self._get_column_name(pattern['col1'], table_name)
            col2 = self._get_column_name(pattern['col2'], table_name)
            if col1 and col2:
                return f'"{col1}" {pattern["op"]} "{col2}"'

        elif pattern['type'] == 'scalar':
            col = self._get_column_name(pattern['col'], table_name)
            if col:
                return f'"{col}" {pattern["op"]} {pattern["value"]}'

        elif pattern['type'] == 'cross_sheet':
            # Cross-sheet references are handled in apply_formula_to_column with special logic
            # This should not be reached, but if it is, return the column name
            target_table = pattern['sheet'].lower().replace(' ', '_')
            col = self._get_column_name(pattern['col'], target_table)
            if col:
                return f'"{col}"'
            return 'NULL'

        elif pattern['type'] == 'if':
            # IF(D2>100, D2*1.1, D2)
            if 'val' in pattern and 'result_val' in pattern:
                col1 = self._get_column_name(pattern['col1'], table_name)
                col_result = self._get_column_name(pattern['result_col'], table_name)
                col_else = self._get_column_name(pattern['else_col'], table_name)
                if col1 and col_result and col_else:
                    return (f'CASE WHEN "{col1}" {pattern["op"]} {pattern["val"]} '
                            f'THEN "{col_result}" {pattern["result_op"]} {pattern["result_val"]} '
                            f'ELSE "{col_else}" END')
            # IF(A2>B2, A2, B2)
            elif 'col2' in pattern:
                col1 = self._get_column_name(pattern['col1'], table_name)
                col2 = self._get_column_name(pattern['col2'], table_name)
                col_result = self._get_column_name(pattern['result_col'], table_name)
                col_else = self._get_column_name(pattern['else_col'], table_name)
                if col1 and col2 and col_result and col_else:
                    return f'CASE WHEN "{col1}" {pattern["op"]} "{col2}" THEN "{col_result}" ELSE "{col_else}" END'

        elif pattern['type'] == 'len_comparison_braced':
            # LEN({TaxID})=13
            header = pattern['header']
            col = self._get_column_by_header(header, table_name)
            if col:
                return f'LENGTH("{col}") {pattern["op"]} {pattern["value"]}'
            return 'NULL'

        elif pattern['type'] == 'len_comparison_range':
            # LEN(D:D)=13
            col = self._get_column_name(pattern['col'], table_name)
            if col:
                return f'LENGTH("{col}") {pattern["op"]} {pattern["value"]}'
            return 'NULL'

        # For complex formulas, use full SQL conversion
        return self.excel_to_sql(formula, table_name.replace('_', ' ').title()).replace('SELECT ', '')

    # ========================================================================
    # RECALCULATION
    # ========================================================================

    def recalculate_all(self) -> None:
        """
        Recalculate all stored formulas.

        Example:
            evaluator.recalculate_all()  # Recalculate all sheets
        """
        tables_to_recalc = list(self.formulas.keys())

        for table_name in tables_to_recalc:
            formulas = self.formulas.get(table_name, {})
            sheet_title = table_name.replace('_', ' ').title()

            for target_column, formula in formulas.items():
                self.apply_formula_to_column(formula, sheet_title, target_column)

    def get_formulas(self) -> Dict[str, Dict[str, str]]:
        """
        Get all stored formulas.

        Returns:
            Dictionary of formulas: {sheet_name: {column: formula}}

        Example:
            formulas = evaluator.get_formulas()
            # {'Sheet1': {'c': '=A2+B2', 'd': '=C2*2'}}
        """
        # Convert table names back to sheet names
        result = {}
        for table_name, formulas in self.formulas.items():
            sheet_name_clean = table_name.replace('_', ' ').title()
            result[sheet_name_clean] = formulas
        return result
