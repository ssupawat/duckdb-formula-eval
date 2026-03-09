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
"""

import re
import duckdb
import pandas as pd
from typing import Any, Dict, List, Tuple, Optional, Union


class FormulaEvaluator:
    """Evaluate Excel formulas using pure DuckDB SQL approach (POC)."""

    def __init__(self, conn: duckdb.DuckDBPyConnection, sheets_data: Dict[str, pd.DataFrame]):
        """
        Initialize the evaluator.

        Args:
            conn: DuckDB connection with registered tables
            sheets_data: Dictionary mapping sheet names to pandas DataFrames
        """
        self.conn = conn
        self.sheets_data = sheets_data
        self.last_sql = None  # Store last generated SQL for debugging

        # Formula storage for recalculation
        # Structure: {sheet_name: {formula_id: {"formula": "...", "target_column": "...", ...}}}
        self.stored_formulas: Dict[str, Dict[str, Dict[str, Any]]] = {}

    # ========================================================================
    # PURE SQL CONVERSION METHODS
    # ========================================================================

    def excel_to_sql(self, formula: str, sheet_name: str, row_ctx: Dict[str, float] = None) -> str:
        """
        Convert Excel formula to pure DuckDB SQL.

        Conversion pipeline order (critical):
        1. String literals (double quotes) → SQL (single quotes)
        2. VLOOKUP → SQL subqueries
        3. Aggregates → SQL subqueries
        4. IF → CASE expressions
        5. Cell references → scalar values
        6. Operators → SQL operators

        Args:
            formula: Excel formula (with or without leading =)
            sheet_name: Name of the sheet containing the formula
            row_ctx: Optional row context for cell references

        Returns:
            SQL SELECT statement that evaluates the formula
        """
        # Remove leading = and whitespace
        expr = formula.lstrip('=').strip()

        # Step 1: VLOOKUP → SQL subqueries (before string literal conversion!)
        expr = self._convert_vlookup_to_sql(expr, sheet_name)

        # Step 2: Convert string literals
        expr = self._convert_string_literals(expr)

        # Step 3: Aggregates → SQL subqueries
        expr = self._convert_aggregates_to_sql(expr, sheet_name)

        # Step 4: IF → CASE expressions
        expr = self._convert_if_to_sql(expr)

        # Step 5: Cell references → scalar values
        if row_ctx:
            expr = self._substitute_cell_references(expr, row_ctx)

        # Step 6: Operators → SQL operators
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
        df = self.sheets_data.get(table_name)
        if df is None:
            return formula

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
            lambda m: '0',  # Empty criteria matches no cells
            formula
        )

        # Handle SUMIF(C:C,'',D:D) pattern - empty criteria with single quotes
        formula = re.sub(
            r"SUMIF\(([A-Z]):([A-Z]),\s*'',\s*([A-Z]):([A-Z])\)",
            lambda m: '0',  # Empty criteria matches no cells
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
            if target_table not in self.sheets_data:
                return '0'

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

        return {'type': 'complex'}

    def _evaluate_vectorized(self, formula: str, sheet_name: str, pattern: Dict[str, Any]) -> pd.Series:
        """Evaluate simple formulas on entire column using vectorized SQL."""
        table_name = sheet_name.lower().replace(' ', '_')
        df = self.sheets_data.get(table_name)
        if df is None:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        if pattern['type'] == 'simple':
            col1 = self._get_column_name(pattern['col1'], table_name)
            col2 = self._get_column_name(pattern['col2'], table_name)
            if col1 and col2:
                sql = f'SELECT "{col1}" {pattern["op"]} "{col2}" FROM {table_name}'
                return self.conn.execute(sql).fetchdf().iloc[:, 0]

        elif pattern['type'] == 'scalar':
            col = self._get_column_name(pattern['col'], table_name)
            if col:
                sql = f'SELECT "{col}" {pattern["op"]} {pattern["value"]} FROM {table_name}'
                return self.conn.execute(sql).fetchdf().iloc[:, 0]

        raise ValueError(f"Unsupported pattern: {pattern}")

    def _get_column_name(self, col_letter: str, sheet_name: str) -> Optional[str]:
        """Map Excel column letter to actual column name."""
        df = self.sheets_data.get(sheet_name)
        if df is None:
            return None

        col_idx = ord(col_letter.upper()) - ord('A')
        if 0 <= col_idx < len(df.columns):
            return df.columns[col_idx]
        return None

    # ========================================================================
    # MAIN EVALUATION ENTRY POINT
    # ========================================================================

    def evaluate_formula(self, formula: str, sheet_name: str, row_ctx: Dict[str, float] = None) -> Union[float, str]:
        """
        Evaluate an Excel formula.

        Args:
            formula: Excel formula (e.g., "=SUM(D:D)", "=D1*1.1")
            sheet_name: Name of the sheet containing the formula
            row_ctx: Optional row context for cell references

        Returns:
            Formula result as float or string
        """
        # Check for vectorized patterns first (only if no row_ctx)
        pattern = self._parse_formula_pattern(formula)
        if row_ctx is None and pattern['type'] in ['simple', 'scalar', 'cross_sheet']:
            try:
                return self._evaluate_vectorized(formula, sheet_name, pattern).iloc[0]
            except Exception:
                pass  # Fall through to SQL evaluation

        # Convert to SQL and execute
        sql = self.excel_to_sql(formula, sheet_name, row_ctx)
        self.last_sql = sql  # For debugging

        result = self.conn.execute(sql).fetchdf().iloc[0, 0]

        if pd.isna(result):
            return 0.0
        elif isinstance(result, str):
            return result
        else:
            return float(result)

    # ========================================================================
    # PERSISTENCE METHODS (POC - simplified)
    # ========================================================================

    def apply_formula_to_column(
        self,
        formula: str,
        sheet_name: str,
        target_column: str,
        context_column: str = None,
        store_formula: bool = True
    ) -> None:
        """
        Apply a formula to all rows in a target column.

        This evaluates the formula for each row and stores the results in the target column.
        If the target column doesn't exist, it will be created.

        Args:
            formula: Excel formula (e.g., "=D2*1.1", "=IF(D2>100, D2*1.1, D2)")
            sheet_name: Name of the sheet
            target_column: Name of column to store results (creates if doesn't exist)
            context_column: Optional column to use for row context (e.g., "D" for D2 references)

        Example:
            # Create a new "bonus" column with 10% increase
            evaluator.apply_formula_to_column('=D2*1.1', 'sheet1', 'bonus', context_column='amount')
        """
        table_name = sheet_name.lower().replace(' ', '_')
        df = self.sheets_data.get(table_name)
        if df is None:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        # Check if target column exists
        if target_column not in df.columns:
            df[target_column] = None

        # Evaluate formula for each row
        results = []
        for idx, row in df.iterrows():
            # Build row context from context_column if specified
            row_ctx = {}
            if context_column and context_column in df.columns:
                # Map context column to Excel cell reference
                col_idx = df.columns.get_loc(context_column)
                col_letter = chr(ord('A') + col_idx)
                cell_ref = f"{col_letter}{idx + 2}"
                row_ctx[cell_ref] = row[context_column]

            # Evaluate formula for this row
            try:
                result = self.evaluate_formula(formula, sheet_name, row_ctx)
                results.append(result)
            except Exception as e:
                results.append(None)

        # Update DataFrame with results
        df[target_column] = results

        # Recreate table with new data
        self._recreate_table_from_dataframe(sheet_name, df)

        # Store formula for recalculation
        if store_formula:
            formula_id = f"{target_column}_formula"
            if table_name not in self.stored_formulas:
                self.stored_formulas[table_name] = {}
            self.stored_formulas[table_name][formula_id] = {
                "formula": formula,
                "target_column": target_column
            }

    def _recreate_table_from_dataframe(self, sheet_name: str, df: pd.DataFrame) -> None:
        """Recreate a DuckDB table from a DataFrame."""
        table_name = sheet_name.lower().replace(' ', '_')
        temp_table = f"{table_name}_new"

        # Create new table with the DataFrame data
        df.to_sql(
            name=temp_table,
            con=self.conn,
            index=False,
            if_exists='replace'
        )

        # Drop old table and rename new one
        self.conn.execute(f'DROP TABLE IF EXISTS {table_name}')
        self.conn.execute(f'ALTER TABLE {temp_table} RENAME TO {table_name}')

        # Update the cached DataFrame
        self.sheets_data[table_name] = df

    def store_formula_at_cell(
        self,
        formula: str,
        sheet_name: str,
        row: int,
        col: str
    ) -> None:
        """
        Store a formula at a specific cell location (Excel-style).

        Args:
            formula: Excel formula (e.g., "=SUM(A:A)")
            sheet_name: Name of the sheet
            row: Row number (1-indexed, like Excel)
            col: Column letter (e.g., "A", "D")

        Example:
            evaluator.store_formula_at_cell('=SUM(A:A)', 'sheet1', row=1, col='F')
        """
        table_name = sheet_name.lower().replace(' ', '_')
        cell_ref = f"{col}{row}"

        if table_name not in self.stored_formulas:
            self.stored_formulas[table_name] = {}

        self.stored_formulas[table_name][cell_ref] = {
            "formula": formula,
            "row": row,
            "col": col
        }

    # ========================================================================
    # RECALCULATION (POC - simple recalculate all)
    # ========================================================================

    def recalculate_all(self, sheet_name: str = None) -> None:
        """
        Recalculate all formulas for a sheet.

        Args:
            sheet_name: Name of the sheet (if None, recalculate all sheets)
        """
        tables_to_recalc = []
        if sheet_name:
            table_name = sheet_name.lower().replace(' ', '_')
            if table_name in self.stored_formulas:
                tables_to_recalc.append(table_name)
        else:
            tables_to_recalc = list(self.stored_formulas.keys())

        for table_name in tables_to_recalc:
            formulas = self.stored_formulas.get(table_name, {})
            sheet_name = table_name.replace('_', ' ').title()

            for formula_id, formula_data in formulas.items():
                formula = formula_data["formula"]
                target_column = formula_data.get("target_column")

                if target_column:
                    # Column formula - reapply to entire column
                    self.apply_formula_to_column(
                        formula,
                        sheet_name,
                        target_column,
                        store_formula=False  # Already stored
                    )

    def get_stored_formulas(self, sheet_name: str = None) -> Dict[str, Dict]:
        """
        Get all stored formulas.

        Args:
            sheet_name: Optional sheet name (if None, return all)

        Returns:
            Dictionary of stored formulas
        """
        if sheet_name:
            table_name = sheet_name.lower().replace(' ', '_')
            return {sheet_name: self.stored_formulas.get(table_name, {})}
        else:
            result = {}
            for table_name, formulas in self.stored_formulas.items():
                sheet_name_clean = table_name.replace('_', ' ').title()
                result[sheet_name_clean] = formulas
            return result
