#!/usr/bin/env python3
"""
DuckDB Excel Formula Evaluator

Evaluates Excel formulas using DuckDB for high-performance SQL-based evaluation.
Supports aggregates, IF statements, VLOOKUP, and nested formulas.

Supported formula types:
- Pure aggregates: SUM, AVERAGE, MAX, MIN, COUNTIF, SUMIF
- Pure scalar: arithmetic, IF statements
- Nested: aggregate inside IF, IF with aggregate conditions
- Arithmetic on aggregates: SUM(D:D)*0.1
- Cross-sheet VLOOKUP
"""

import re
import duckdb
import pandas as pd
import numexpr
from typing import Any, Dict, List, Tuple, Optional


class FormulaEvaluator:
    """Evaluate Excel formulas using hybrid SQL + numexpr approach."""

    def __init__(self, conn: duckdb.DuckDBPyConnection, sheets_data: Dict[str, pd.DataFrame]):
        """
        Initialize the evaluator.

        Args:
            conn: DuckDB connection with registered tables
            sheets_data: Dictionary mapping sheet names to pandas DataFrames
        """
        self.conn = conn
        self.sheets_data = sheets_data

    def evaluate_formula(self, formula: str, sheet_name: str, row_ctx: Dict[str, float] = None) -> float | str:
        """
        Evaluate a formula using two-phase approach:
        Phase 1: Extract and compute all aggregates using SQL
        Phase 2: Substitute aggregates and evaluate scalar expression

        Args:
            formula: Excel formula (e.g., "=SUM(A:A)", "=IF(D1>80, D1*1.1, D1*0.9)")
            sheet_name: Name of the sheet containing the formula
            row_ctx: Optional row context for cell references (e.g., {"D1": 100.0})

        Returns:
            Either a numeric value or a string (for VLOOKUP results)
        """
        simplified = self._resolve_aggregates(formula, sheet_name)
        return self._evaluate_scalar(simplified, row_ctx or {})

    def _resolve_aggregates(self, formula: str, sheet_name: str) -> str:
        """Phase 1: Replace all aggregate functions with their computed values."""
        result = formula

        # Patterns with their SQL equivalents (most specific first)
        patterns = [
            # SUMIF/COUNTIF with criteria - handle both quote styles
            (r'SUMIF\(([A-Z]+):([A-Z]+),[\'"]([^\'"]+)[\'"],([A-Z]+):([A-Z]+)\)', self._sumif),
            (r'COUNTIF\(([A-Z]+):([A-Z]+),[\'"]([^\'"]+)[\'"]\)', self._countif),

            # Basic aggregates
            (r'SUM\(([A-Z]+):([A-Z]+)\)', self._sum),
            (r'AVERAGE\(([A-Z]+):([A-Z]+)\)', self._average),
            (r'MAX\(([A-Z]+):([A-Z]+)\)', self._max),
            (r'MIN\(([A-Z]+):([A-Z]+)\)', self._min),
            (r'COUNT\(([A-Z]+):([A-Z]+)\)', self._count),
        ]

        # Iteratively resolve aggregates (handles nesting)
        max_iterations = 10
        for _ in range(max_iterations):
            prev_result = result
            for pattern, handler in patterns:
                def replace_match(m):
                    return str(handler(m, sheet_name))

                result = re.sub(pattern, replace_match, result, flags=re.IGNORECASE)

            if result == prev_result:
                break  # No more changes

        return result

    def _get_column_name(self, col_letter: str, sheet_name: str) -> str:
        """Map Excel column letter to actual column name in the sheet."""
        df = self.sheets_data.get(sheet_name.lower().replace(' ', '_'))
        if df is None:
            return f'col_{col_letter.lower()}'

        # Column letter to index (A=0, B=1, etc.)
        col_idx = ord(col_letter.upper()) - ord('A')
        if col_idx < len(df.columns):
            return df.columns[col_idx]
        return f'col_{col_letter.lower()}'

    def _sum(self, m: re.Match, sheet_name: str) -> float:
        col = self._get_column_name(m.group(1), sheet_name)
        result = self.conn.execute(f"SELECT COALESCE(SUM(\"{col}\"), 0) FROM {sheet_name}").fetchone()[0]
        return float(result)

    def _average(self, m: re.Match, sheet_name: str) -> float:
        col = self._get_column_name(m.group(1), sheet_name)
        result = self.conn.execute(f"SELECT COALESCE(AVG(\"{col}\"), 0) FROM {sheet_name}").fetchone()[0]
        return float(result)

    def _max(self, m: re.Match, sheet_name: str) -> float:
        col = self._get_column_name(m.group(1), sheet_name)
        result = self.conn.execute(f"SELECT COALESCE(MAX(\"{col}\"), 0) FROM {sheet_name}").fetchone()[0]
        return float(result)

    def _min(self, m: re.Match, sheet_name: str) -> float:
        col = self._get_column_name(m.group(1), sheet_name)
        result = self.conn.execute(f"SELECT COALESCE(MIN(\"{col}\"), 0) FROM {sheet_name}").fetchone()[0]
        return float(result)

    def _count(self, m: re.Match, sheet_name: str) -> int:
        col = self._get_column_name(m.group(1), sheet_name)
        result = self.conn.execute(f"SELECT COUNT(*) FROM {sheet_name} WHERE \"{col}\" IS NOT NULL").fetchone()[0]
        return int(result)

    def _sumif(self, m: re.Match, sheet_name: str) -> float:
        criteria_col = self._get_column_name(m.group(1), sheet_name)
        criteria_val = m.group(3)  # Already stripped by regex
        sum_col = self._get_column_name(m.group(4), sheet_name)

        # Criteria is already extracted without quotes by regex
        # Just use it directly with quotes for string comparison
        result = self.conn.execute(
            f'SELECT COALESCE(SUM(\"{sum_col}\"), 0) FROM {sheet_name} WHERE \"{criteria_col}\" = \'{criteria_val}\''
        ).fetchone()[0]

        return float(result)

    def _countif(self, m: re.Match, sheet_name: str) -> int:
        col = self._get_column_name(m.group(1), sheet_name)
        criteria_val = m.group(3)  # Already stripped by regex (group 3 for COUNTIF)

        # Criteria is already extracted without quotes by regex
        result = self.conn.execute(
            f'SELECT COUNT(*) FROM {sheet_name} WHERE \"{col}\" = \'{criteria_val}\''
        ).fetchone()[0]

        return int(result)

    def _evaluate_scalar(self, formula: str, row_ctx: Dict[str, float]) -> float | str:
        """
        Phase 2: Evaluate scalar expression using numexpr for performance and security.
        Handles IF, arithmetic operations, VLOOKUP, and cell references.

        Returns either a numeric value or a string (for VLOOKUP results).
        """
        # Remove leading = if present
        expr = formula.lstrip('=').strip()

        # Substitute cell references with values from row_ctx
        for ref, value in row_ctx.items():
            # Match whole cell references like D1, not partial matches
            # Quote string values to preserve them in the expression
            if isinstance(value, str):
                replacement = f'"{value}"'
            else:
                replacement = str(value)
            expr = re.sub(rf'\b{re.escape(ref)}\b', replacement, expr)

        # Convert Excel's = to Python's == for equality comparisons
        # Pattern: value="value" or value=value -> value=="value" or value==value
        # But avoid converting == that's already there, or in function calls
        expr = re.sub(r'(?<=[\w"\'])=', '==', expr)

        # Handle VLOOKUP: VLOOKUP(lookup_value, table_array, col_index, range_lookup)
        # Check if formula is just a VLOOKUP (standalone, not part of larger expression)
        if re.match(r'^VLOOKUP\(', expr):
            expr, vlookup_result = self._process_vlookup(expr)
            if vlookup_result is not None:
                return vlookup_result
        else:
            # VLOOKUP is part of larger expression - process it
            expr, vlookup_result = self._process_vlookup(expr)

        # Handle IF statements: IF(condition, true_value, false_value)
        # This converts to Python ternary: (true_val if condition else false_val)
        expr = self._process_if_statements(expr)

        # Convert Python ternary to numexpr where clause format
        # Python: (true_val if condition else false_val)
        # numexpr: where(condition, true_val, false_val)
        expr = self._convert_to_numexpr(expr)

        # Evaluate the expression using numexpr (faster and safer than eval)
        try:
            result = numexpr.evaluate(expr, local_dict={})
            # numexpr returns a numpy array, extract scalar value
            if hasattr(result, 'item'):
                return float(result.item())
            return float(result)
        except Exception as e:
            raise ValueError(f"Failed to evaluate expression '{expr}': {e}")

    def _process_if_statements(self, expr: str) -> str:
        """Process nested IF statements."""
        # Pattern: IF(condition, true_val, false_val)
        if_pattern = r'IF\(([^,]+),([^,]+),([^)]+)\)'

        def replace_if(m):
            condition = m.group(1).strip()
            true_val = m.group(2).strip()
            false_val = m.group(3).strip()

            # Recursively process nested IFs
            condition = self._process_if_statements(condition)
            true_val = self._process_if_statements(true_val)
            false_val = self._process_if_statements(false_val)

            # Return Python ternary expression
            return f"({true_val} if {condition} else {false_val})"

        # Process from innermost to outermost (reverse the matches)
        max_iterations = 10
        for _ in range(max_iterations):
            prev = expr
            expr = re.sub(if_pattern, replace_if, expr)
            if expr == prev:
                break

        return expr

    def _process_vlookup(self, expr: str) -> tuple:
        """Process VLOOKUP function. Returns (new_expr, vlookup_result)."""
        # Pattern: VLOOKUP(lookup_value, Sheet2!A:B, col_index, range_lookup)
        vlookup_pattern = r'VLOOKUP\(([^,]+),([A-Za-z0-9_]+)!([A-Z]+):([A-Z]+),(\d+),([01]+)\)'
        vlookup_result = None

        def replace_vlookup(m):
            nonlocal vlookup_result
            lookup_val = m.group(1).strip().strip('"\'')  # Remove quotes
            sheet = m.group(2).lower().replace(' ', '_')
            start_col = m.group(3)
            end_col = m.group(4)
            col_offset = int(m.group(5)) - 1  # 1-based to 0-based
            range_lookup = m.group(6) == '1'

            # Get the source sheet data
            df = self.sheets_data.get(sheet)
            if df is None:
                return "0"

            # Find the lookup column (first column in the range)
            lookup_col_letter = start_col
            lookup_col_idx = ord(lookup_col_letter.upper()) - ord('A')

            if lookup_col_idx >= len(df.columns):
                return "0"

            lookup_col_name = df.columns[lookup_col_idx]

            # Find the return column
            return_col_idx = lookup_col_idx + col_offset
            if return_col_idx >= len(df.columns):
                return "0"

            return_col_name = df.columns[return_col_idx]

            # Perform the lookup
            try:
                # Try to convert lookup_val to number for numeric comparison
                try:
                    numeric_lookup = float(lookup_val)
                    # Try numeric comparison first
                    if range_lookup:
                        # Approximate match (find closest)
                        result = df.iloc[(df[lookup_col_name].astype(float) - numeric_lookup).abs().argsort()[:1]]
                    else:
                        # Exact match
                        result = df[df[lookup_col_name].astype(float) == numeric_lookup]
                except ValueError:
                    # String comparison
                    if range_lookup:
                        # Approximate match (not well-defined for strings, use exact)
                        result = df[df[lookup_col_name].astype(str) == lookup_val]
                    else:
                        # Exact match
                        result = df[df[lookup_col_name].astype(str) == lookup_val]

                if len(result) > 0 and return_col_name in result.columns:
                    val = result[return_col_name].iloc[0]
                    vlookup_result = val  # Store for return
                    # Return the actual value
                    return str(val)
                return "0"
            except Exception:
                return "0"

        new_expr = re.sub(vlookup_pattern, replace_vlookup, expr)
        return (new_expr, vlookup_result)

    def _convert_to_numexpr(self, expr: str) -> str:
        """Convert Python ternary expression to numexpr format.

        Python ternary: (true_val if condition else false_val)
        numexpr format: where(condition, true_val, false_val)
        """
        # Pattern: (true_val if condition else false_val)
        # Note: Python ternary puts condition in the middle
        # We need to capture: (value_if_true if condition else value_if_false)
        pattern = r'\(([^:]+?)\s+if\s+([^:]+?)\s+else\s+([^)]+?)\)'

        def convert_match(m):
            true_val = m.group(1).strip()
            condition = m.group(2).strip()
            false_val = m.group(3).strip()
            return f'where({condition}, {true_val}, {false_val})'

        # Process nested ternary (innermost first)
        max_iterations = 5
        for _ in range(max_iterations):
            prev = expr
            expr = re.sub(pattern, convert_match, expr)
            if expr == prev:
                break

        return expr
