"""
Parameter validator for Excel operations.

Validates operation parameters to ensure they are safe, well-formed,
and compatible with the target Excel operations.
"""

from typing import Dict, List, Any, Optional, Union
from dataclasses import dataclass
import re
from datetime import datetime


@dataclass
class ValidationResult:
    """Result of parameter validation."""
    valid: bool
    errors: List[str]
    warnings: List[str]
    sanitized_parameters: Dict[str, Any]


class ParameterValidator:
    """Validates and sanitizes operation parameters."""
    
    def __init__(self):
        """Initialize parameter validator with validation rules."""
        self._required_parameters = {
            'insert_row': ['sheet_name', 'data'],
            'insert_column': ['sheet_name', 'column_name'],
            'update_cells': ['sheet_name', 'range'],
            'delete_rows': ['sheet_name', 'conditions'],
            'filter_data': ['sheet_name', 'conditions'],
            'create_chart': ['sheet_name', 'data_range', 'chart_type'],
            'aggregate_data': ['sheet_name', 'columns', 'operation'],
        }
        
        self._parameter_types = {
            'sheet_name': str,
            'data': (list, dict, str),
            'range': str,
            'conditions': (list, dict, str),
            'chart_type': str,
            'columns': (list, str),
            'operation': str,
            'max_rows': int,
            'max_columns': int,
            'position': int,
            'chart_id': (str, int),
            'axis': str,
            'amount': (int, float),
        }
        
        self._valid_chart_types = [
            'bar', 'line', 'pie', 'scatter', 'area', 'column',
            'histogram', 'box', 'bubble', 'radar'
        ]
        
        self._valid_aggregation_operations = [
            'sum', 'avg', 'average', 'count', 'max', 'min',
            'median', 'std', 'var', 'first', 'last'
        ]
        
        self._dangerous_strings = [
            '=', 'INDIRECT', 'HYPERLINK', 'EXEC', 'SHELL',
            'CALL', 'SYSTEM', 'IMPORT', 'EVAL'
        ]
    
    def validate_parameters(self, operation: str, 
                          parameters: Dict[str, Any]) -> ValidationResult:
        """
        Validate parameters for a specific operation.
        
        Args:
            operation: The operation name
            parameters: Parameters to validate
            
        Returns:
            ValidationResult with validation status and sanitized parameters
        """
        errors = []
        warnings = []
        sanitized = parameters.copy()
        
        # Check required parameters
        required_errors = self._check_required_parameters(operation, parameters)
        errors.extend(required_errors)
        
        # Validate parameter types
        type_errors, type_warnings, sanitized = self._validate_parameter_types(
            parameters, sanitized
        )
        errors.extend(type_errors)
        warnings.extend(type_warnings)
        
        # Validate parameter values
        value_errors, value_warnings, sanitized = self._validate_parameter_values(
            operation, sanitized
        )
        errors.extend(value_errors)
        warnings.extend(value_warnings)
        
        # Check for security issues
        security_errors, sanitized = self._check_security_issues(sanitized)
        errors.extend(security_errors)
        
        # Validate operation-specific parameters
        op_errors, op_warnings, sanitized = self._validate_operation_specific(
            operation, sanitized
        )
        errors.extend(op_errors)
        warnings.extend(op_warnings)
        
        return ValidationResult(
            valid=len(errors) == 0,
            errors=errors,
            warnings=warnings,
            sanitized_parameters=sanitized
        )
    
    def _check_required_parameters(self, operation: str, 
                                 parameters: Dict[str, Any]) -> List[str]:
        """Check if all required parameters are present."""
        errors = []
        required = self._required_parameters.get(operation, [])
        
        for param in required:
            if param not in parameters or parameters[param] is None:
                errors.append(f"Missing required parameter: {param}")
        
        return errors
    
    def _validate_parameter_types(self, parameters: Dict[str, Any], 
                                sanitized: Dict[str, Any]) -> tuple[List[str], List[str], Dict[str, Any]]:
        """Validate and convert parameter types."""
        errors = []
        warnings = []
        
        for param, value in parameters.items():
            if param in self._parameter_types:
                expected_type = self._parameter_types[param]
                
                # Handle multiple allowed types
                if isinstance(expected_type, tuple):
                    if not isinstance(value, expected_type):
                        # Try to convert to the first type in the tuple
                        try:
                            sanitized[param] = expected_type[0](value)
                            warnings.append(f"Converted {param} to {expected_type[0].__name__}")
                        except (ValueError, TypeError):
                            errors.append(
                                f"Parameter {param} must be one of types: "
                                f"{', '.join(t.__name__ for t in expected_type)}"
                            )
                else:
                    if not isinstance(value, expected_type):
                        # Try to convert
                        try:
                            sanitized[param] = expected_type(value)
                            warnings.append(f"Converted {param} to {expected_type.__name__}")
                        except (ValueError, TypeError):
                            errors.append(f"Parameter {param} must be of type {expected_type.__name__}")
        
        return errors, warnings, sanitized
    
    def _validate_parameter_values(self, operation: str, 
                                 sanitized: Dict[str, Any]) -> tuple[List[str], List[str], Dict[str, Any]]:
        """Validate parameter values and ranges."""
        errors = []
        warnings = []
        
        # Validate numeric ranges
        if 'max_rows' in sanitized:
            max_rows = sanitized['max_rows']
            if max_rows < 1:
                errors.append("max_rows must be at least 1")
            elif max_rows > 50:
                errors.append("max_rows cannot exceed 50 for safety")
        
        if 'max_columns' in sanitized:
            max_columns = sanitized['max_columns']
            if max_columns < 1:
                errors.append("max_columns must be at least 1")
            elif max_columns > 20:
                errors.append("max_columns cannot exceed 20 for safety")
        
        if 'position' in sanitized:
            position = sanitized['position']
            if position < 1:
                errors.append("position must be at least 1")
        
        # Validate string parameters
        if 'sheet_name' in sanitized:
            sheet_name = sanitized['sheet_name']
            if not sheet_name or not isinstance(sheet_name, str):
                errors.append("sheet_name must be a non-empty string")
            elif len(sheet_name) > 31:  # Excel sheet name limit
                errors.append("sheet_name cannot exceed 31 characters")
            elif any(char in sheet_name for char in ['/', '\\', '?', '*', '[', ']']):
                errors.append("sheet_name contains invalid characters")
        
        # Validate range parameters
        if 'range' in sanitized:
            range_str = sanitized['range']
            if not self._is_valid_excel_range(range_str):
                errors.append(f"Invalid Excel range format: {range_str}")
        
        # Validate chart types
        if 'chart_type' in sanitized:
            chart_type = sanitized['chart_type'].lower()
            if chart_type not in self._valid_chart_types:
                errors.append(
                    f"Invalid chart type: {chart_type}. "
                    f"Valid types: {', '.join(self._valid_chart_types)}"
                )
            sanitized['chart_type'] = chart_type
        
        # Validate aggregation operations
        if 'operation' in sanitized and operation == 'aggregate_data':
            agg_op = sanitized['operation'].lower()
            if agg_op not in self._valid_aggregation_operations:
                errors.append(
                    f"Invalid aggregation operation: {agg_op}. "
                    f"Valid operations: {', '.join(self._valid_aggregation_operations)}"
                )
            sanitized['operation'] = agg_op
        
        return errors, warnings, sanitized
    
    def _check_security_issues(self, sanitized: Dict[str, Any]) -> tuple[List[str], Dict[str, Any]]:
        """Check for potential security issues in parameters."""
        errors = []
        
        # Check all string parameters for dangerous content
        for param, value in sanitized.items():
            if isinstance(value, str):
                for dangerous in self._dangerous_strings:
                    if dangerous in value.upper():
                        errors.append(
                            f"Parameter {param} contains potentially dangerous content: {dangerous}"
                        )
                        break
        
        # Check data parameter for formula injection
        if 'data' in sanitized:
            data = sanitized['data']
            if isinstance(data, list):
                for item in data:
                    if isinstance(item, str) and item.startswith('='):
                        errors.append("Data contains formula that could be dangerous")
                        break
            elif isinstance(data, str) and data.startswith('='):
                errors.append("Data contains formula that could be dangerous")
        
        # Check conditions for SQL injection-like patterns
        if 'conditions' in sanitized:
            conditions = str(sanitized['conditions']).upper()
            dangerous_sql = ['DROP', 'DELETE', 'TRUNCATE', 'ALTER', 'CREATE', 'INSERT']
            for dangerous in dangerous_sql:
                if dangerous in conditions:
                    errors.append(f"Conditions contain potentially dangerous keyword: {dangerous}")
        
        return errors, sanitized
    
    def _validate_operation_specific(self, operation: str, 
                                   sanitized: Dict[str, Any]) -> tuple[List[str], List[str], Dict[str, Any]]:
        """Validate parameters specific to each operation type."""
        errors = []
        warnings = []
        
        if operation == 'insert_row':
            errors_ir, warnings_ir, sanitized = self._validate_insert_row(sanitized)
            errors.extend(errors_ir)
            warnings.extend(warnings_ir)
        
        elif operation == 'delete_rows':
            errors_dr, warnings_dr, sanitized = self._validate_delete_rows(sanitized)
            errors.extend(errors_dr)
            warnings.extend(warnings_dr)
        
        elif operation == 'create_chart':
            errors_cc, warnings_cc, sanitized = self._validate_create_chart(sanitized)
            errors.extend(errors_cc)
            warnings.extend(warnings_cc)
        
        elif operation == 'filter_data':
            errors_fd, warnings_fd, sanitized = self._validate_filter_data(sanitized)
            errors.extend(errors_fd)
            warnings.extend(warnings_fd)
        
        return errors, warnings, sanitized
    
    def _validate_insert_row(self, sanitized: Dict[str, Any]) -> tuple[List[str], List[str], Dict[str, Any]]:
        """Validate insert_row specific parameters."""
        errors = []
        warnings = []
        
        if 'data' in sanitized:
            data = sanitized['data']
            if isinstance(data, list) and len(data) > 100:
                errors.append("Cannot insert more than 100 columns in a single row")
            elif isinstance(data, dict) and len(data) > 100:
                errors.append("Cannot insert more than 100 key-value pairs in a single row")
        
        return errors, warnings, sanitized
    
    def _validate_delete_rows(self, sanitized: Dict[str, Any]) -> tuple[List[str], List[str], Dict[str, Any]]:
        """Validate delete_rows specific parameters."""
        errors = []
        warnings = []
        
        if 'conditions' in sanitized:
            conditions = sanitized['conditions']
            if not conditions or conditions == ['*']:
                errors.append("Delete operations must have specific conditions")
        
        # Ensure max_rows is set for delete operations
        if 'max_rows' not in sanitized:
            sanitized['max_rows'] = 10  # Conservative default
            warnings.append("Added default max_rows limit of 10 for delete operation")
        
        return errors, warnings, sanitized
    
    def _validate_create_chart(self, sanitized: Dict[str, Any]) -> tuple[List[str], List[str], Dict[str, Any]]:
        """Validate create_chart specific parameters."""
        errors = []
        warnings = []
        
        if 'data_range' in sanitized:
            data_range = sanitized['data_range']
            if not self._is_valid_excel_range(data_range):
                errors.append(f"Invalid data range for chart: {data_range}")
        
        return errors, warnings, sanitized
    
    def _validate_filter_data(self, sanitized: Dict[str, Any]) -> tuple[List[str], List[str], Dict[str, Any]]:
        """Validate filter_data specific parameters."""
        errors = []
        warnings = []
        
        if 'conditions' in sanitized:
            conditions = sanitized['conditions']
            if isinstance(conditions, list) and len(conditions) > 10:
                warnings.append("Large number of filter conditions may impact performance")
        
        return errors, warnings, sanitized
    
    def _is_valid_excel_range(self, range_str: str) -> bool:
        """Check if a string is a valid Excel range format."""
        if not range_str or not isinstance(range_str, str):
            return False
        
        range_str = range_str.upper().strip()
        
        # Single cell (A1, B2, etc.)
        if re.match(r'^[A-Z]+\d+$', range_str):
            return True
        
        # Range (A1:B2, A:A, 1:1, etc.)
        if ':' in range_str:
            parts = range_str.split(':')
            if len(parts) == 2:
                start, end = parts
                # Cell range (A1:B2)
                if re.match(r'^[A-Z]+\d+$', start) and re.match(r'^[A-Z]+\d+$', end):
                    return True
                # Column range (A:B)
                if re.match(r'^[A-Z]+$', start) and re.match(r'^[A-Z]+$', end):
                    return True
                # Row range (1:5)
                if re.match(r'^\d+$', start) and re.match(r'^\d+$', end):
                    return True
        
        return False
    
    def get_parameter_help(self, operation: str) -> Dict[str, str]:
        """Get help text for parameters of a specific operation."""
        help_text = {
            'insert_row': {
                'sheet_name': 'Name of the Excel sheet to insert into',
                'data': 'List or dictionary of data to insert',
                'position': 'Row position to insert at (optional, defaults to end)'
            },
            'delete_rows': {
                'sheet_name': 'Name of the Excel sheet',
                'conditions': 'Conditions to match rows for deletion',
                'max_rows': 'Maximum number of rows to delete (safety limit)'
            },
            'create_chart': {
                'sheet_name': 'Name of the Excel sheet',
                'data_range': 'Excel range containing the data (e.g., A1:B10)',
                'chart_type': f'Type of chart: {", ".join(self._valid_chart_types)}',
                'title': 'Chart title (optional)'
            }
        }
        
        return help_text.get(operation, {})
    
    def sanitize_string_parameter(self, value: str) -> str:
        """Sanitize a string parameter to remove dangerous content."""
        if not isinstance(value, str):
            return str(value)
        
        # Remove dangerous characters and patterns
        sanitized = value
        for dangerous in self._dangerous_strings:
            sanitized = sanitized.replace(dangerous, '')
        
        # Limit length
        if len(sanitized) > 1000:
            sanitized = sanitized[:1000]
        
        return sanitized.strip()