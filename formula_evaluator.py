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
from typing import Any, Dict, List, Tuple, Optional, Union


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

    def _parse_formula_pattern(self, formula: str) -> Dict[str, Any]:
        """
        Parse a formula to detect patterns that can be vectorized.

        Returns:
            {
                'type': 'simple' | 'scalar' | 'cross_sheet' | 'complex',
                'pattern': e.g., 'A+B' or 'A*2' or 'Sheet1!A',
                'columns': list of column letters involved,
                'source_sheet': for cross_sheet patterns
            }
        """
        # Remove leading = and whitespace
        formula = formula.lstrip('=').strip()

        # Simple arithmetic between two columns: A2+B2, A2*C2, etc.
        simple_arithmetic = re.match(r'^([A-Z])\d+\s*([+\-*/])\s*([A-Z])\d+$', formula)
        if simple_arithmetic:
            return {
                'type': 'simple',
                'pattern': f'{simple_arithmetic.group(1)} {simple_arithmetic.group(2)} {simple_arithmetic.group(3)}',
                'columns': [simple_arithmetic.group(1), simple_arithmetic.group(3)]
            }

        # Simple scalar operation on a column: A2*2, B2/10, etc.
        simple_scalar = re.match(r'^([A-Z])\d+\s*([+\-*/])\s*(\d+(?:\.\d+)?)$', formula)
        if simple_scalar:
            return {
                'type': 'scalar',
                'pattern': f'{simple_scalar.group(1)} {simple_scalar.group(2)} {simple_scalar.group(3)}',
                'columns': [simple_scalar.group(1)]
            }

        # Cross-sheet reference: Sheet1!A2
        cross_sheet = re.match(r'^([A-Za-z0-9_]+)!([A-Z])\d+$', formula)
        if cross_sheet:
            return {
                'type': 'cross_sheet',
                'source_sheet': cross_sheet.group(1),
                'column': cross_sheet.group(2)
            }

        # Complex formulas require the original two-phase approach
        return {'type': 'complex'}

    def _evaluate_vectorized(self, formula: str, sheet_name: str, pattern: Dict[str, Any]) -> pd.Series:
        """
        Evaluate simple formulas using vectorized SQL for entire columns.

        Returns a pandas Series with the results for all rows.
        """
        table_name = sheet_name.lower().replace(' ', '_')
        df = self.sheets_data.get(table_name)
        if df is None:
            raise ValueError(f"Sheet '{sheet_name}' not found")

        # Build column mapping: Excel letter (A, B, C) to actual column name
        col_map = {chr(ord('A') + i): df.columns[i] for i in range(len(df.columns))}

        if pattern['type'] == 'simple':
            # Two-column arithmetic: A+B, A-C, etc.
            parts = pattern['pattern'].split()
            col1 = col_map[parts[0]]
            op = parts[1]
            col2 = col_map[parts[2]]

            sql_expr = f'"{col1}" {op} "{col2}"'
            result_df = self.conn.execute(f'SELECT {sql_expr} FROM {table_name}').fetchdf()
            return result_df.iloc[:, 0]

        elif pattern['type'] == 'scalar':
            # Column with scalar: A*2, B/10, etc.
            parts = pattern['pattern'].split()
            col = col_map[parts[0]]
            op = parts[1]
            scalar_val = parts[2]

            sql_expr = f'"{col}" {op} {scalar_val}'
            result_df = self.conn.execute(f'SELECT {sql_expr} FROM {table_name}').fetchdf()
            return result_df.iloc[:, 0]

        elif pattern['type'] == 'cross_sheet':
            # Reference to another sheet: Sheet1!A
            source_table = pattern['source_sheet'].lower().replace(' ', '_')
            source_df = self.sheets_data.get(source_table)
            if source_df is None:
                raise ValueError(f"Source sheet '{pattern['source_sheet']}' not found")

            source_col_map = {chr(ord('A') + i): source_df.columns[i] for i in range(len(source_df.columns))}
            source_col = source_col_map[pattern['column']]

            result_df = self.conn.execute(f'SELECT "{source_col}" FROM {source_table}').fetchdf()
            return result_df.iloc[:, 0]

        raise ValueError(f"Unsupported pattern type: {pattern['type']}")

    def evaluate_formula(self, formula: str, sheet_name: str, row_ctx: Dict[str, float] = None) -> Union[float, str]:
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
            # SUMIF/COUNTIF with criteria - handle both quote styles and unquoted criteria
            (r'SUMIF\(([A-Z]+):([A-Z]+),\s*(?:\"([^\"]+)\"|\'([^\']+)\'|([^,)]+)),([A-Z]+):([A-Z]+)\)', self._sumif),
            (r'COUNTIF\(([A-Z]+):([A-Z]+),\s*(?:\"([^\"]+)\"|\'([^\']+)\'|([^,)]+))\)', self._countif),

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

    def _get_column_name(self, col_letter: str, sheet_name: str) -> Optional[str]:
        """Map Excel column letter to actual column name in the sheet. Returns None if column doesn't exist."""
        df = self.sheets_data.get(sheet_name.lower().replace(' ', '_'))
        if df is None:
            return None

        # Column letter to index (A=0, B=1, etc.)
        col_idx = ord(col_letter.upper()) - ord('A')
        if col_idx < len(df.columns):
            return df.columns[col_idx]
        return None

    def _sum(self, m: re.Match, sheet_name: str) -> float:
        col = self._get_column_name(m.group(1), sheet_name)
        if col is None:
            return 0.0
        result = self.conn.execute(f"SELECT COALESCE(SUM(\"{col}\"), 0) FROM {sheet_name}").fetchone()[0]
        return float(result)

    def _average(self, m: re.Match, sheet_name: str) -> float:
        col = self._get_column_name(m.group(1), sheet_name)
        if col is None:
            return 0.0
        result = self.conn.execute(f"SELECT COALESCE(AVG(\"{col}\"), 0) FROM {sheet_name}").fetchone()[0]
        return float(result)

    def _max(self, m: re.Match, sheet_name: str) -> float:
        col = self._get_column_name(m.group(1), sheet_name)
        if col is None:
            return 0.0
        result = self.conn.execute(f"SELECT COALESCE(MAX(\"{col}\"), 0) FROM {sheet_name}").fetchone()[0]
        return float(result)

    def _min(self, m: re.Match, sheet_name: str) -> float:
        col = self._get_column_name(m.group(1), sheet_name)
        if col is None:
            return 0.0
        result = self.conn.execute(f"SELECT COALESCE(MIN(\"{col}\"), 0) FROM {sheet_name}").fetchone()[0]
        return float(result)

    def _count(self, m: re.Match, sheet_name: str) -> float:
        col = self._get_column_name(m.group(1), sheet_name)
        if col is None:
            return 0.0
        result = self.conn.execute(f"SELECT COUNT(*) FROM {sheet_name} WHERE \"{col}\" IS NOT NULL").fetchone()[0]
        return float(result)

    def _sumif(self, m: re.Match, sheet_name: str) -> float:
        criteria_col = self._get_column_name(m.group(1), sheet_name)
        if criteria_col is None:
            return 0.0
        # Extract criteria_val from the correct group (3, 4, or 5)
        # Groups 3 and 4 are quoted (already stripped by regex), group 5 is unquoted
        criteria_val_raw = (m.group(3) or m.group(4) or m.group(5) or "").strip()
        sum_col = self._get_column_name(m.group(6), sheet_name)
        if sum_col is None:
            return 0.0

        # Handle empty criteria
        if not criteria_val_raw or criteria_val_raw == '""':
            # Empty criteria matches NULL/empty cells
            where_clause = f'"{criteria_col}" IS NULL'
        else:
            # Parse criteria for operator and value
            operator_match = re.match(r'([<>=!]+)\s*(.*)', criteria_val_raw)
            if operator_match:
                op = operator_match.group(1)
                val_str = operator_match.group(2).strip()
                # Convert operators
                if op == '=':
                    op = '=='
                elif op == '<>':
                    op = '!='

                try:
                    # Try to convert value to a number for comparison
                    val = float(val_str)
                    where_clause = f'"{criteria_col}" {op} {val}'
                except ValueError:
                    # If not a number, treat as string literal (e.g., "=Text")
                    where_clause = f'"{criteria_col}" {op} \'{val_str}\''
            else:
                # No operator, assume exact match
                where_clause = f'"{criteria_col}" = \'{criteria_val_raw}\''

        result = self.conn.execute(
            f'SELECT COALESCE(SUM(\"{sum_col}\"), 0) FROM {sheet_name} WHERE {where_clause}'
        ).fetchone()[0]

        return float(result)

    def _countif(self, m: re.Match, sheet_name: str) -> float:
        col = self._get_column_name(m.group(1), sheet_name)
        if col is None:
            return 0.0
        # Extract criteria_val from the correct group (3, 4, or 5)
        criteria_val_raw = (m.group(3) or m.group(4) or m.group(5)).strip()

        # Parse criteria for operator and value
        operator_match = re.match(r'([<>=!]+)\s*(.*)', criteria_val_raw)
        if operator_match:
            op = operator_match.group(1)
            val_str = operator_match.group(2).strip()
            if op == '=':
                op = '=='
            elif op == '<>':
                op = '!='

            try:
                val = float(val_str)
                where_clause = f'"{col}" {op} {val}'
            except ValueError:
                where_clause = f'"{col}" {op} \'{val_str}\''
        else:
            # No operator, assume exact match
            where_clause = f'"{col}" = \'{criteria_val_raw}\''

        result = self.conn.execute(
            f'SELECT COUNT(*) FROM {sheet_name} WHERE {where_clause}'
        ).fetchone()[0]

        return float(result)

    def _evaluate_scalar(self, formula: str, row_ctx: Dict[str, float]) -> Union[float, str]:
        """
        Phase 2: Evaluate scalar expression using numexpr for performance and security.
        Handles IF, arithmetic operations, VLOOKUP, and cell references.

        Returns either a numeric value or a string (for VLOOKUP results).
        """
        # Remove leading = if present
        expr = formula.lstrip('=').strip()

        # Convert Excel's <> to Python's != before other processing
        expr = expr.replace('<>', '!=')

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
        # Only convert if not already ==, and avoid converting in function calls
        expr = re.sub(r'(?<=[\w"\'])=(?!=)', '==', expr)

        # Handle VLOOKUP: VLOOKUP(lookup_value, table_array, col_index, range_lookup)
        # Check if formula is just a VLOOKUP (standalone, not part of larger expression)
        if re.match(r'^VLOOKUP\(', expr):
            expr, vlookup_result = self._process_vlookup(expr)
            if vlookup_result is not None:
                return vlookup_result
        else:
            # VLOOKUP is part of larger expression - process it
            expr, vlookup_result = self._process_vlookup(expr)

        # Check if this is an IF statement that returns strings (can't use numexpr)
        if self._has_string_if_result(expr):
            return self._evaluate_string_if(expr)

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

    def _has_string_if_result(self, expr: str) -> bool:
        """Check if the expression contains IF statements with string results."""
        # Look for IF statements with quoted string values
        if_pattern = r'IF\(([^,]+),([^,]+),([^)]+)\)'
        matches = re.findall(if_pattern, expr)
        for condition, true_val, false_val in matches:
            # Check if either branch is a quoted string
            if ('"' in true_val or '"' in false_val or "'" in true_val or "'" in false_val):
                return True
        return False

    def _evaluate_string_if(self, expr: str) -> Union[float, str]:
        """Evaluate IF statements with string results using Python eval."""
        # Remove outer quotes for eval, but keep inner string quotes
        expr = expr.strip()

        # Process IF statements
        if_pattern = r'IF\(([^,]+),([^,]+),([^)]+)\)'

        def replace_if(m):
            condition = m.group(1).strip()
            true_val = m.group(2).strip()
            false_val = m.group(3).strip()

            # Recursively process nested IFs
            condition = self._evaluate_string_if(condition) if 'IF(' in condition else condition
            true_val = self._evaluate_string_if(true_val) if 'IF(' in true_val else true_val
            false_val = self._evaluate_string_if(false_val) if 'IF(' in false_val else false_val

            # Build Python eval expression
            # Remove quotes from string values for eval, keep them for strings
            def clean_val(v):
                v = v.strip()
                if (v.startswith('"') and v.endswith('"')) or (v.startswith("'") and v.endswith("'")):
                    return v  # Keep quotes for strings
                return v  # Return numbers/expressions as-is

            true_val = clean_val(true_val)
            false_val = clean_val(false_val)

            return f"({true_val} if {condition} else {false_val})"

        # Iteratively replace IF statements (from innermost first)
        max_iterations = 10
        for _ in range(max_iterations):
            prev = expr
            expr = re.sub(if_pattern, replace_if, expr)
            if expr == prev:
                break

        # Evaluate using Python's eval (safe here since we control the input)
        try:
            result = eval(expr, {"__builtins__": {}}, {})
            return result
        except Exception as e:
            raise ValueError(f"Failed to evaluate string IF expression '{expr}': {e}")

    def _process_if_statements(self, expr: str) -> str:
        """Process nested IF statements. Only for numeric results."""
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
            # Find innermost IF first by using non-greedy matching
            matches = list(re.finditer(r'IF\(([^(),]+(?:\([^()]*\))?[^(),]*),([^(),]+(?:\([^()]*\))?[^(),]*),([^()]+(?:\([^()]*\))?[^()]*)\)', expr))
            if not matches:
                break
            # Process from end (innermost) to start
            for m in reversed(matches):
                original = m.group(0)
                replacement = replace_if(m)
                expr = expr.replace(original, replacement, 1)
                break  # Process one replacement, then re-scan

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
                vlookup_result = 0
                return "0"

            # Find the lookup column (first column in the range)
            lookup_col_letter = start_col
            lookup_col_idx = ord(lookup_col_letter.upper()) - ord('A')

            if lookup_col_idx >= len(df.columns):
                vlookup_result = 0
                return "0"

            lookup_col_name = df.columns[lookup_col_idx]

            # Find the return column
            return_col_idx = lookup_col_idx + col_offset
            if return_col_idx >= len(df.columns):
                vlookup_result = 0
                return "0"

            return_col_name = df.columns[return_col_idx]

            # Perform the lookup
            try:
                # Try to convert lookup_val to number for numeric comparison
                try:
                    numeric_lookup = float(lookup_val)
                    lookup_series = df[lookup_col_name].astype(float)

                    if range_lookup:
                        # Approximate match: find largest value <= lookup_val
                        # Filter to values <= lookup_val, then take max
                        valid_values = lookup_series[lookup_series <= numeric_lookup]
                        if len(valid_values) > 0:
                            idx = valid_values.idxmax()
                            result = df.loc[[idx]]
                        else:
                            result = pd.DataFrame()
                    else:
                        # Exact match
                        result = df[lookup_series == numeric_lookup]
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
                vlookup_result = 0
                return "0"
            except Exception:
                vlookup_result = 0
                return "0"

        new_expr = re.sub(vlookup_pattern, replace_vlookup, expr)
        return (new_expr, vlookup_result)

    def _convert_to_numexpr(self, expr: str) -> str:
        """Convert Python ternary expression to numexpr format.

        Python ternary: (true_val if condition else false_val)
        numexpr format: where(condition, true_val, false_val)
        """
        # Process from innermost to outermost
        max_iterations = 10
        for _ in range(max_iterations):
            prev = expr
            # Find innermost ternary: (value_if_true if condition else value_if_false)
            inner_pattern = r'\(([^():]+)\s+if\s+([^():]+)\s+else\s+([^():]+)\)'

            def convert_match(m):
                true_val = m.group(1).strip()
                condition = m.group(2).strip()
                false_val = m.group(3).strip()
                return f'where({condition}, {true_val}, {false_val})'

            expr = re.sub(inner_pattern, convert_match, expr)
            if expr == prev:
                break

        return expr
