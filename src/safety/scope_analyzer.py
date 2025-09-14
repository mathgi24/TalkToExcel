"""
Operation scope analyzer for Excel operations.

Analyzes and enforces limits on operation scope to prevent
mass operations that could damage spreadsheet data.
"""

from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass
import re


@dataclass
class ScopeAnalysis:
    """Result of scope analysis for an operation."""
    estimated_rows: int
    estimated_columns: int
    within_limits: bool
    violations: List[str]
    suggested_alternatives: List[str]


class ScopeAnalyzer:
    """Analyzes and enforces operation scope limits."""
    
    def __init__(self, max_rows: int = 50, max_columns: int = 20):
        """
        Initialize scope analyzer with limits.
        
        Args:
            max_rows: Maximum number of rows that can be affected
            max_columns: Maximum number of columns that can be affected
        """
        self.max_rows = max_rows
        self.max_columns = max_columns
    
    def analyze_scope(self, operation: str, parameters: Dict[str, Any], 
                     sheet_info: Optional[Dict[str, Any]] = None) -> ScopeAnalysis:
        """
        Analyze the scope of an operation.
        
        Args:
            operation: The operation name
            parameters: Operation parameters
            sheet_info: Information about the target sheet (rows, columns, etc.)
            
        Returns:
            ScopeAnalysis with estimated impact and limit violations
        """
        estimated_rows = 0
        estimated_columns = 0
        violations = []
        suggested_alternatives = []
        
        # Analyze based on operation type
        if operation in ['delete_rows', 'update_cells', 'clear_data']:
            estimated_rows, estimated_columns = self._analyze_destructive_operation(
                parameters, sheet_info
            )
        elif operation in ['insert_row', 'insert_column']:
            estimated_rows, estimated_columns = self._analyze_insertion_operation(
                parameters
            )
        elif operation in ['filter_data', 'sort_data', 'aggregate_data']:
            estimated_rows, estimated_columns = self._analyze_query_operation(
                parameters, sheet_info
            )
        else:
            # Default analysis for unknown operations
            estimated_rows, estimated_columns = self._analyze_generic_operation(
                parameters, sheet_info
            )
        
        # Check for violations
        if estimated_rows > self.max_rows:
            violations.append(f"Operation affects {estimated_rows} rows (limit: {self.max_rows})")
            suggested_alternatives.append(
                f"Consider breaking this into smaller operations of {self.max_rows} rows or less"
            )
            suggested_alternatives.append("Add more specific conditions to limit the scope")
        
        if estimated_columns > self.max_columns:
            violations.append(f"Operation affects {estimated_columns} columns (limit: {self.max_columns})")
            suggested_alternatives.append(
                f"Consider specifying only the columns you need (max {self.max_columns})"
            )
        
        within_limits = len(violations) == 0
        
        return ScopeAnalysis(
            estimated_rows=estimated_rows,
            estimated_columns=estimated_columns,
            within_limits=within_limits,
            violations=violations,
            suggested_alternatives=suggested_alternatives
        )
    
    def _analyze_destructive_operation(self, parameters: Dict[str, Any], 
                                     sheet_info: Optional[Dict[str, Any]]) -> Tuple[int, int]:
        """Analyze scope for destructive operations like delete or update."""
        rows = 0
        columns = 0
        
        # Check explicit row/column limits
        if 'max_rows' in parameters:
            rows = parameters['max_rows']
        
        if 'max_columns' in parameters:
            columns = parameters['max_columns']
        
        # Analyze range parameter
        if 'range' in parameters:
            range_rows, range_cols = self._parse_range(parameters['range'])
            rows = max(rows, range_rows)
            columns = max(columns, range_cols)
        
        # Analyze conditions - if no conditions, assume affects all data
        if 'conditions' in parameters:
            conditions = parameters.get('conditions', [])
            if not conditions or conditions == ['*'] or 'all' in str(conditions).lower():
                # No limiting conditions - could affect entire sheet
                if sheet_info:
                    rows = sheet_info.get('row_count', 1000)  # Conservative estimate
                    columns = sheet_info.get('column_count', 50)
                else:
                    rows = 1000  # Conservative estimate for unknown sheets
                    columns = 50
        
        return rows, columns
    
    def _analyze_insertion_operation(self, parameters: Dict[str, Any]) -> Tuple[int, int]:
        """Analyze scope for insertion operations."""
        rows = 1  # Default: single row insertion
        columns = 1  # Default: single column insertion
        
        # Check if inserting multiple rows/columns
        if 'count' in parameters:
            count = parameters.get('count', 1)
            if 'row' in parameters.get('type', '').lower():
                rows = count
            else:
                columns = count
        
        # Check data parameter for bulk insertions
        if 'data' in parameters:
            data = parameters['data']
            if isinstance(data, list):
                if len(data) > 0 and isinstance(data[0], list):
                    # 2D array - rows x columns
                    rows = len(data)
                    columns = len(data[0]) if data else 1
                else:
                    # 1D array - assume single row/column
                    rows = 1
                    columns = len(data)
        
        return rows, columns
    
    def _analyze_query_operation(self, parameters: Dict[str, Any], 
                               sheet_info: Optional[Dict[str, Any]]) -> Tuple[int, int]:
        """Analyze scope for query operations (usually read-only)."""
        # Query operations are generally safe but we still track scope
        rows = 0
        columns = 0
        
        if sheet_info:
            # Queries typically read data, so estimate based on sheet size
            rows = min(sheet_info.get('row_count', 100), 100)  # Cap at 100 for queries
            
            if 'columns' in parameters:
                columns = len(parameters['columns'])
            else:
                columns = min(sheet_info.get('column_count', 10), 10)  # Cap at 10
        
        return rows, columns
    
    def _analyze_generic_operation(self, parameters: Dict[str, Any], 
                                 sheet_info: Optional[Dict[str, Any]]) -> Tuple[int, int]:
        """Generic scope analysis for unknown operations."""
        rows = 1
        columns = 1
        
        # Look for common parameters that indicate scope
        scope_indicators = {
            'range': self._parse_range,
            'max_rows': lambda x: (x, 1),
            'max_columns': lambda x: (1, x),
            'row_count': lambda x: (x, 1),
            'column_count': lambda x: (1, x)
        }
        
        for param, analyzer in scope_indicators.items():
            if param in parameters:
                param_rows, param_cols = analyzer(parameters[param])
                rows = max(rows, param_rows)
                columns = max(columns, param_cols)
        
        return rows, columns
    
    def _parse_range(self, range_str: str) -> Tuple[int, int]:
        """
        Parse Excel range string to estimate affected rows and columns.
        
        Args:
            range_str: Excel range like "A1:B10" or "A:A" or "1:5"
            
        Returns:
            Tuple of (estimated_rows, estimated_columns)
        """
        if not range_str or not isinstance(range_str, str):
            return 1, 1
        
        range_str = range_str.upper().strip()
        
        # Handle entire column references (A:A, B:Z, etc.)
        if re.match(r'^[A-Z]+:[A-Z]+$', range_str):
            start_col, end_col = range_str.split(':')
            col_count = ord(end_col) - ord(start_col) + 1
            return 1000, col_count  # Assume large number of rows for entire columns
        
        # Handle entire row references (1:1, 1:10, etc.)
        if re.match(r'^\d+:\d+$', range_str):
            start_row, end_row = map(int, range_str.split(':'))
            row_count = end_row - start_row + 1
            return row_count, 50  # Assume many columns for entire rows
        
        # Handle cell range references (A1:B10)
        if ':' in range_str:
            try:
                start_cell, end_cell = range_str.split(':')
                start_row, start_col = self._parse_cell_reference(start_cell)
                end_row, end_col = self._parse_cell_reference(end_cell)
                
                row_count = abs(end_row - start_row) + 1
                col_count = abs(ord(end_col) - ord(start_col)) + 1
                
                return row_count, col_count
            except (ValueError, IndexError):
                # If parsing fails, assume single cell
                return 1, 1
        
        # Single cell reference
        return 1, 1
    
    def _parse_cell_reference(self, cell_ref: str) -> Tuple[int, str]:
        """
        Parse cell reference like "A1" into row number and column letter.
        
        Returns:
            Tuple of (row_number, column_letter)
        """
        match = re.match(r'^([A-Z]+)(\d+)$', cell_ref.upper())
        if not match:
            raise ValueError(f"Invalid cell reference: {cell_ref}")
        
        col_letters = match.group(1)
        row_number = int(match.group(2))
        
        return row_number, col_letters
    
    def get_scope_summary(self, analysis: ScopeAnalysis) -> str:
        """Get human-readable summary of scope analysis."""
        if analysis.within_limits:
            return (f"Operation scope: {analysis.estimated_rows} rows, "
                   f"{analysis.estimated_columns} columns (within limits)")
        else:
            violations_text = "; ".join(analysis.violations)
            return f"Operation scope exceeds limits: {violations_text}"
    
    def suggest_scope_reduction(self, analysis: ScopeAnalysis) -> List[str]:
        """Suggest ways to reduce operation scope."""
        suggestions = []
        
        if analysis.estimated_rows > self.max_rows:
            suggestions.append(
                f"Add conditions to limit affected rows to {self.max_rows} or fewer"
            )
            suggestions.append(
                "Consider processing data in batches"
            )
        
        if analysis.estimated_columns > self.max_columns:
            suggestions.append(
                f"Specify only the columns you need (max {self.max_columns})"
            )
            suggestions.append(
                "Use column names instead of ranges when possible"
            )
        
        suggestions.extend(analysis.suggested_alternatives)
        
        return list(set(suggestions))  # Remove duplicates