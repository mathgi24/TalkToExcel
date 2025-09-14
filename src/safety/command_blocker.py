"""
Command blocker for dangerous Excel operations.

Blocks dangerous operations and provides helpful error messages
with suggestions for safer alternatives.
"""

from typing import Dict, List, Any, Optional
from dataclasses import dataclass
import re


@dataclass
class BlockResult:
    """Result of command blocking analysis."""
    blocked: bool
    reason: str
    error_message: str
    suggested_alternatives: List[str]


class CommandBlocker:
    """Blocks dangerous operations with helpful error messages."""
    
    def __init__(self):
        """Initialize command blocker with dangerous operation patterns."""
        self._blocked_operations = {
            'format_all': {
                'reason': 'Mass formatting operations are not allowed',
                'message': 'This operation is not allowed as it would affect the entire spreadsheet formatting. Please specify a specific range or cells to format.',
                'alternatives': [
                    'Specify a cell range like "A1:B10" instead of "all"',
                    'Format individual columns or rows',
                    'Use conditional formatting for specific criteria'
                ]
            },
            'delete_all': {
                'reason': 'Mass deletion operations are not permitted',
                'message': 'Mass deletion operations are not permitted for safety. Please specify exact rows, columns, or conditions for deletion.',
                'alternatives': [
                    'Delete specific rows by row number',
                    'Delete based on specific conditions',
                    'Clear content instead of deleting structure'
                ]
            },
            'clear_all': {
                'reason': 'Mass clear operations are not allowed',
                'message': 'Clearing all data is not permitted for safety. Please specify the exact range or conditions for clearing data.',
                'alternatives': [
                    'Clear specific cell ranges',
                    'Clear based on conditions',
                    'Clear individual sheets instead of entire workbook'
                ]
            },
            'replace_all': {
                'reason': 'Mass replace operations are dangerous',
                'message': 'Mass replace operations could cause unintended data changes. Please specify conditions or ranges for replacement.',
                'alternatives': [
                    'Replace within specific columns',
                    'Use find and replace with conditions',
                    'Preview changes before applying'
                ]
            }
        }
        
        self._dangerous_patterns = [
            # Mass operation patterns
            r'\b(format|delete|clear|remove|replace)\s+(all|everything|entire|whole)\b',
            r'\b(all|everything|entire|whole)\s+(format|delete|clear|remove|replace)\b',
            
            # Entire spreadsheet patterns
            r'\bentire\s+(spreadsheet|workbook|file|sheet)\b',
            r'\bwhole\s+(spreadsheet|workbook|file|sheet)\b',
            r'\bcomplete\s+(spreadsheet|workbook|file|sheet)\b',
            
            # Dangerous range patterns
            r'\bA:Z\b',  # Entire spreadsheet range
            r'\b1:1048576\b',  # All rows in Excel
            r'\bA1:XFD1048576\b',  # Entire Excel range
            
            # Dangerous keywords combinations
            r'\bdelete\s+everything\b',
            r'\bclear\s+everything\b',
            r'\bremove\s+all\s+data\b',
            r'\bformat\s+entire\b',
        ]
        
        self._scope_limits = {
            'max_rows': 50,
            'max_columns': 20,
            'max_cells': 1000
        }
    
    def check_command(self, operation: str, parameters: Dict[str, Any], 
                     command_text: str = "") -> BlockResult:
        """
        Check if a command should be blocked.
        
        Args:
            operation: The operation name
            parameters: Operation parameters
            command_text: Original natural language command
            
        Returns:
            BlockResult indicating if command is blocked and why
        """
        # Check for explicitly blocked operations
        if operation in self._blocked_operations:
            block_info = self._blocked_operations[operation]
            return BlockResult(
                blocked=True,
                reason=block_info['reason'],
                error_message=block_info['message'],
                suggested_alternatives=block_info['alternatives']
            )
        
        # Check for dangerous patterns in command text
        pattern_result = self._check_dangerous_patterns(command_text)
        if pattern_result.blocked:
            return pattern_result
        
        # Check for scope violations
        scope_result = self._check_scope_violations(parameters)
        if scope_result.blocked:
            return scope_result
        
        # Check for parameter-based blocking
        param_result = self._check_dangerous_parameters(parameters)
        if param_result.blocked:
            return param_result
        
        # Command is not blocked
        return BlockResult(
            blocked=False,
            reason="",
            error_message="",
            suggested_alternatives=[]
        )
    
    def _check_dangerous_patterns(self, command_text: str) -> BlockResult:
        """Check command text for dangerous patterns."""
        if not command_text:
            return BlockResult(False, "", "", [])
        
        command_lower = command_text.lower()
        
        for pattern in self._dangerous_patterns:
            if re.search(pattern, command_lower, re.IGNORECASE):
                return BlockResult(
                    blocked=True,
                    reason="Command contains dangerous mass operation pattern",
                    error_message=(
                        "This command appears to be a mass operation that could affect "
                        "large amounts of data. For safety, please specify exact ranges, "
                        "conditions, or limits for your operation."
                    ),
                    suggested_alternatives=[
                        "Specify exact cell ranges (e.g., A1:B10)",
                        "Add conditions to limit the operation scope",
                        "Break the operation into smaller, specific tasks",
                        "Use preview mode to see what would be affected"
                    ]
                )
        
        return BlockResult(False, "", "", [])
    
    def _check_scope_violations(self, parameters: Dict[str, Any]) -> BlockResult:
        """Check for scope violations that should block the operation."""
        violations = []
        alternatives = []
        
        # Check row limits
        max_rows = parameters.get('max_rows', 0)
        if max_rows > self._scope_limits['max_rows']:
            violations.append(f"Operation affects {max_rows} rows (limit: {self._scope_limits['max_rows']})")
            alternatives.append(f"Limit operation to {self._scope_limits['max_rows']} rows or fewer")
        
        # Check column limits
        max_columns = parameters.get('max_columns', 0)
        if max_columns > self._scope_limits['max_columns']:
            violations.append(f"Operation affects {max_columns} columns (limit: {self._scope_limits['max_columns']})")
            alternatives.append(f"Limit operation to {self._scope_limits['max_columns']} columns or fewer")
        
        # Check for unlimited operations
        conditions = parameters.get('conditions', [])
        if not conditions or conditions == ['*'] or 'all' in str(conditions).lower():
            if parameters.get('operation_type') in ['delete', 'update', 'clear']:
                violations.append("Operation has no limiting conditions and could affect all data")
                alternatives.append("Add specific conditions to limit which data is affected")
        
        if violations:
            return BlockResult(
                blocked=True,
                reason="Operation scope exceeds safety limits",
                error_message=(
                    f"Operation blocked for safety: {'; '.join(violations)}. "
                    "Please reduce the scope of your operation."
                ),
                suggested_alternatives=alternatives
            )
        
        return BlockResult(False, "", "", [])
    
    def _check_dangerous_parameters(self, parameters: Dict[str, Any]) -> BlockResult:
        """Check for dangerous parameter combinations."""
        dangerous_params = []
        alternatives = []
        
        # Check for dangerous range parameters
        range_param = parameters.get('range', '')
        if isinstance(range_param, str):
            dangerous_ranges = ['A:Z', '1:1048576', 'A1:XFD1048576', 'entire', 'all']
            if any(dangerous in range_param.upper() for dangerous in dangerous_ranges):
                dangerous_params.append("Range parameter covers entire spreadsheet")
                alternatives.append("Specify a smaller, specific range")
        
        # Check for dangerous operation types
        operation_type = parameters.get('operation_type', '').lower()
        target = parameters.get('target', '').lower()
        
        if operation_type in ['delete', 'clear', 'format'] and target in ['all', 'everything', 'entire']:
            dangerous_params.append(f"Dangerous combination: {operation_type} {target}")
            alternatives.append(f"Specify exact items to {operation_type}")
        
        # Check for formula injection attempts
        if 'formula' in parameters or 'expression' in parameters:
            formula = parameters.get('formula', '') or parameters.get('expression', '')
            if isinstance(formula, str) and ('=' in formula or 'INDIRECT' in formula.upper()):
                dangerous_params.append("Formula parameter may contain dangerous expressions")
                alternatives.append("Use simple values instead of complex formulas")
        
        if dangerous_params:
            return BlockResult(
                blocked=True,
                reason="Dangerous parameter combination detected",
                error_message=(
                    f"Operation blocked due to dangerous parameters: {'; '.join(dangerous_params)}. "
                    "Please modify your request to be more specific and safer."
                ),
                suggested_alternatives=alternatives
            )
        
        return BlockResult(False, "", "", [])
    
    def get_blocked_operations(self) -> List[str]:
        """Get list of all blocked operation names."""
        return list(self._blocked_operations.keys())
    
    def add_blocked_operation(self, operation: str, reason: str, 
                            message: str, alternatives: List[str]):
        """
        Add a new blocked operation.
        
        Args:
            operation: Operation name to block
            reason: Short reason for blocking
            message: Detailed error message for users
            alternatives: List of suggested alternatives
        """
        self._blocked_operations[operation] = {
            'reason': reason,
            'message': message,
            'alternatives': alternatives
        }
    
    def is_operation_blocked(self, operation: str) -> bool:
        """Check if a specific operation is blocked."""
        return operation in self._blocked_operations
    
    def get_safe_alternatives(self, blocked_operation: str) -> List[str]:
        """Get safe alternatives for a blocked operation."""
        if blocked_operation in self._blocked_operations:
            return self._blocked_operations[blocked_operation]['alternatives']
        return []
    
    def format_error_message(self, block_result: BlockResult) -> str:
        """Format a comprehensive error message for blocked operations."""
        if not block_result.blocked:
            return ""
        
        message = f"❌ Operation Blocked: {block_result.error_message}\n"
        
        if block_result.suggested_alternatives:
            message += "\n💡 Suggested alternatives:\n"
            for i, alternative in enumerate(block_result.suggested_alternatives, 1):
                message += f"  {i}. {alternative}\n"
        
        message += "\nFor safety, this system prevents operations that could cause widespread data loss or corruption."
        
        return message