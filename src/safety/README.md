# Safety and Validation System

The safety and validation system provides comprehensive protection for Excel operations by implementing multiple layers of safety checks, risk assessment, and parameter validation.

## Overview

This system prevents dangerous operations that could cause data loss or corruption while providing helpful guidance to users for safer alternatives. It consists of four main components coordinated by a central safety manager.

## Components

### 1. Risk Assessor (`risk_assessor.py`)

Classifies operations into risk levels based on their potential impact:

- **Low Risk**: Read operations, simple queries, single cell updates
- **Medium Risk**: Structural changes, multiple row operations, chart modifications  
- **High Risk**: Destructive operations, large scope operations
- **Dangerous**: Mass operations, operations affecting entire spreadsheets

```python
from src.safety.risk_assessor import RiskAssessor, RiskLevel

assessor = RiskAssessor()
assessment = assessor.assess_operation(
    'delete_rows',
    {'max_rows': 25, 'conditions': ['status = inactive']},
    'delete inactive records'
)

print(f"Risk Level: {assessment.level.value}")
print(f"Blocked: {assessment.blocked}")
print(f"Confirmation Required: {assessment.confirmation_required}")
```

### 2. Scope Analyzer (`scope_analyzer.py`)

Analyzes and enforces limits on operation scope to prevent mass operations:

- **Row Limit**: Maximum 50 rows per operation
- **Column Limit**: Maximum 20 columns per operation
- **Range Analysis**: Parses Excel ranges to estimate impact
- **Violation Detection**: Identifies operations exceeding limits

```python
from src.safety.scope_analyzer import ScopeAnalyzer

analyzer = ScopeAnalyzer(max_rows=50, max_columns=20)
analysis = analyzer.analyze_scope(
    'update_cells',
    {'range': 'A1:B100'},  # Exceeds row limit
    {'row_count': 1000, 'column_count': 30}
)

print(f"Within Limits: {analysis.within_limits}")
print(f"Violations: {analysis.violations}")
```

### 3. Command Blocker (`command_blocker.py`)

Blocks dangerous operations with helpful error messages:

- **Blocked Operations**: `format_all`, `delete_all`, `clear_all`, `replace_all`
- **Pattern Detection**: Identifies dangerous keywords in natural language
- **Parameter Analysis**: Blocks based on dangerous parameter combinations
- **Alternative Suggestions**: Provides safer alternatives for blocked operations

```python
from src.safety.command_blocker import CommandBlocker

blocker = CommandBlocker()
result = blocker.check_command(
    'format_all',
    {'sheet_name': 'Sheet1'},
    'format all cells'
)

if result.blocked:
    print(f"Blocked: {result.error_message}")
    print(f"Alternatives: {result.suggested_alternatives}")
```

### 4. Parameter Validator (`parameter_validator.py`)

Validates and sanitizes operation parameters:

- **Required Parameters**: Ensures all required parameters are present
- **Type Validation**: Validates and converts parameter types
- **Value Validation**: Checks parameter values and ranges
- **Security Checks**: Prevents formula injection and dangerous content
- **Operation-Specific Rules**: Custom validation for each operation type

```python
from src.safety.parameter_validator import ParameterValidator

validator = ParameterValidator()
result = validator.validate_parameters(
    'create_chart',
    {
        'sheet_name': 'Sales',
        'data_range': 'A1:B10',
        'chart_type': 'BAR'  # Will be converted to lowercase
    }
)

print(f"Valid: {result.valid}")
print(f"Sanitized Parameters: {result.sanitized_parameters}")
print(f"Warnings: {result.warnings}")
```

### 5. Safety Manager (`safety_manager.py`)

Coordinates all safety components and provides a unified interface:

```python
from src.safety.safety_manager import SafetyManager

safety_manager = SafetyManager(max_rows=50, max_columns=20)

# Comprehensive safety evaluation
result = safety_manager.evaluate_operation(
    'delete_rows',
    {
        'sheet_name': 'Data',
        'conditions': ['status = inactive'],
        'max_rows': 25
    },
    'delete inactive records'
)

# Check results
print(f"Safe: {result.safe}")
print(f"Blocked: {result.blocked}")
print(f"Confirmation Required: {result.confirmation_required}")

# Get user-friendly summary
print(safety_manager.get_safety_summary(result))

# Generate confirmation prompt if needed
if result.confirmation_required:
    prompt = safety_manager.create_confirmation_prompt(result)
    print(prompt)
```

## Safety Levels and Actions

| Risk Level | Action | Description |
|------------|--------|-------------|
| **Low** | ✅ Allow | Safe operations that don't modify data structure |
| **Medium** | ❓ Confirm | Operations that modify data, require user confirmation |
| **High** | ⚠️ Confirm + Backup | Destructive operations, require confirmation and backup |
| **Dangerous** | 🚫 Block | Mass operations that could cause widespread damage |

## Blocked Operations

The following operations are completely blocked for safety:

- `format_all` - Mass formatting operations
- `delete_all` - Mass deletion operations  
- `clear_all` - Mass clear operations
- `replace_all` - Mass replace operations

### Dangerous Command Patterns

Commands containing these patterns are automatically blocked:

- "format all", "delete everything", "clear entire"
- "entire spreadsheet", "whole workbook"
- Range patterns like "A:Z", "1:1048576"

## Safety Limits

### Scope Limits
- **Maximum Rows**: 50 rows per operation
- **Maximum Columns**: 20 columns per operation
- **Maximum Cells**: 1000 cells per operation

### Parameter Limits
- **Sheet Name**: Maximum 31 characters (Excel limit)
- **String Parameters**: Maximum 1000 characters
- **Array Data**: Maximum 100 items for single operations

## Security Features

### Formula Injection Prevention
- Blocks parameters containing `=`, `INDIRECT`, `HYPERLINK`
- Sanitizes string inputs to remove dangerous content
- Validates data arrays for formula patterns

### Input Sanitization
- Removes dangerous characters from string parameters
- Validates Excel range formats
- Checks for SQL injection-like patterns in conditions

## Integration Example

```python
from src.safety.safety_manager import SafetyManager

class ExcelCommandProcessor:
    def __init__(self):
        self.safety_manager = SafetyManager()
    
    def process_command(self, operation, parameters, command_text):
        # Evaluate safety
        safety_result = self.safety_manager.evaluate_operation(
            operation, parameters, command_text
        )
        
        # Handle different outcomes
        if safety_result.blocked:
            return self.handle_blocked_operation(safety_result)
        elif safety_result.confirmation_required:
            return self.request_user_confirmation(safety_result)
        elif safety_result.safe:
            return self.execute_operation(safety_result)
        else:
            return self.handle_safety_concerns(safety_result)
```

## Error Messages and Suggestions

The safety system provides helpful error messages and suggestions:

### Blocked Operation Example
```
❌ Operation Blocked: Mass deletion operations are not permitted for safety. 
Please specify exact rows, columns, or conditions for deletion.

💡 Suggested alternatives:
  1. Delete specific rows by row number
  2. Delete based on specific conditions  
  3. Clear content instead of deleting structure
```

### Scope Violation Example
```
⚠️ Operation scope exceeds limits: Operation affects 75 rows (limit: 50)

💡 Suggestions:
  1. Add conditions to limit affected rows to 50 or fewer
  2. Consider processing data in batches
  3. Break the operation into smaller, specific tasks
```

## Testing

Comprehensive tests are available in `tests/test_safety_system.py`:

```bash
python -m pytest tests/test_safety_system.py -v
```

Or run the demo:

```bash
python examples/safety_system_demo.py
```

## Configuration

Safety limits can be customized:

```python
# Custom limits
safety_manager = SafetyManager(max_rows=100, max_columns=50)

# Update limits at runtime
safety_manager.update_safety_limits(max_rows=25, max_columns=10)

# Add custom blocked operations
safety_manager.add_custom_blocked_operation(
    'dangerous_custom_op',
    'Custom dangerous operation',
    'This operation is blocked for custom reasons',
    ['Use safe alternative instead']
)
```

## Best Practices

1. **Always use SafetyManager** for operation evaluation
2. **Handle confirmation prompts** appropriately in your UI
3. **Create backups** before high-risk operations
4. **Provide clear error messages** to users
5. **Log safety events** for audit purposes
6. **Test edge cases** with the safety system
7. **Update safety rules** as new operations are added

## Future Enhancements

- Machine learning-based risk assessment
- User-specific safety profiles
- Operation history analysis
- Advanced pattern recognition
- Integration with Excel's built-in safety features