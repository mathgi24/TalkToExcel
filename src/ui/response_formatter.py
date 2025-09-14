"""
Response formatting and display system.

This module provides comprehensive response formatting for different types of
command processing results, ensuring clear and user-friendly output.
"""

import json
from typing import Any, Dict, List, Optional, Union
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from processing.command_processor import ProcessingResult, ProcessingStatus

# Add src directory to Python path for imports
import sys
from pathlib import Path
current_dir = Path(__file__).parent
src_dir = current_dir.parent if current_dir.name != 'src' else current_dir
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))



@dataclass
class FormattingOptions:
    """Options for response formatting."""
    show_timestamps: bool = True
    show_operation_details: bool = False
    max_data_rows: int = 20
    max_column_width: int = 30
    use_colors: bool = True
    compact_mode: bool = False


class ResponseFormatter:
    """
    Formats command processing results for display to users.
    
    Handles different result types and provides clear, readable output
    with appropriate formatting for success, errors, confirmations, etc.
    """
    
    def __init__(self, options: Optional[FormattingOptions] = None):
        """
        Initialize response formatter.
        
        Args:
            options: Formatting options, uses defaults if None
        """
        self.options = options or FormattingOptions()
        
        # Color codes for terminal output
        self.colors = {
            'success': '\033[92m',    # Green
            'error': '\033[91m',      # Red
            'warning': '\033[93m',    # Yellow
            'info': '\033[94m',       # Blue
            'prompt': '\033[95m',     # Magenta
            'reset': '\033[0m',       # Reset
            'bold': '\033[1m',        # Bold
            'dim': '\033[2m'          # Dim
        } if self.options.use_colors else {k: '' for k in ['success', 'error', 'warning', 'info', 'prompt', 'reset', 'bold', 'dim']}
    
    def format_response(self, result: ProcessingResult) -> str:
        """
        Format a processing result for display.
        
        Args:
            result: Processing result to format
            
        Returns:
            Formatted string ready for display
        """
        if result.status == ProcessingStatus.SUCCESS:
            return self._format_success_response(result)
        elif result.status == ProcessingStatus.FAILED:
            return self._format_error_response(result)
        elif result.status == ProcessingStatus.BLOCKED:
            return self._format_blocked_response(result)
        elif result.status == ProcessingStatus.CONFIRMATION_REQUIRED:
            return self._format_confirmation_response(result)
        elif result.status == ProcessingStatus.CLARIFICATION_NEEDED:
            return self._format_clarification_response(result)
        else:
            return self._format_generic_response(result)
    
    def _format_success_response(self, result: ProcessingResult) -> str:
        """Format successful operation response."""
        lines = []
        
        # Success header
        lines.append(f"{self.colors['success']}✅ SUCCESS{self.colors['reset']}")
        
        # Main message
        if result.message:
            lines.append(f"{self.colors['bold']}{result.message}{self.colors['reset']}")
        
        # Data display
        if result.data is not None:
            data_display = self._format_data(result.data)
            if data_display:
                lines.append("")
                lines.append(f"{self.colors['info']}📊 Results:{self.colors['reset']}")
                lines.extend(data_display)
        
        # Operation details
        if result.operation_details and self.options.show_operation_details:
            lines.append("")
            lines.append(f"{self.colors['dim']}Operation: {result.operation_details.get('operation', 'unknown')}{self.colors['reset']}")
        
        # Warnings
        if result.warnings:
            lines.append("")
            lines.append(f"{self.colors['warning']}⚠️ Warnings:{self.colors['reset']}")
            for warning in result.warnings:
                lines.append(f"  • {warning}")
        
        # Timestamp
        if self.options.show_timestamps:
            timestamp = datetime.now().strftime("%H:%M:%S")
            lines.append(f"{self.colors['dim']}[{timestamp}]{self.colors['reset']}")
        
        return "\n".join(lines)
    
    def _format_error_response(self, result: ProcessingResult) -> str:
        """Format error response."""
        lines = []
        
        # Error header
        lines.append(f"{self.colors['error']}❌ ERROR{self.colors['reset']}")
        
        # Error message
        if result.message:
            lines.append(f"{self.colors['bold']}{result.message}{self.colors['reset']}")
        
        # Additional error details
        if result.operation_details:
            error_details = result.operation_details.get('error_details')
            if error_details:
                lines.append("")
                lines.append(f"{self.colors['dim']}Details: {error_details}{self.colors['reset']}")
        
        # Suggestions
        if result.warnings:  # Using warnings field for suggestions
            lines.append("")
            lines.append(f"{self.colors['info']}💡 Suggestions:{self.colors['reset']}")
            for suggestion in result.warnings:
                lines.append(f"  • {suggestion}")
        
        return "\n".join(lines)
    
    def _format_blocked_response(self, result: ProcessingResult) -> str:
        """Format blocked operation response."""
        lines = []
        
        # Blocked header
        lines.append(f"{self.colors['error']}🚫 OPERATION BLOCKED{self.colors['reset']}")
        
        # Block reason
        if result.message:
            lines.append(f"{self.colors['bold']}{result.message}{self.colors['reset']}")
        
        # Safety report
        if result.safety_report:
            lines.append("")
            lines.append(f"{self.colors['warning']}🛡️ Safety Report:{self.colors['reset']}")
            lines.extend(self._format_safety_report(result.safety_report))
        
        # Alternative suggestions
        if result.warnings:
            lines.append("")
            lines.append(f"{self.colors['info']}💡 Alternative approaches:{self.colors['reset']}")
            for suggestion in result.warnings:
                lines.append(f"  • {suggestion}")
        
        return "\n".join(lines)
    
    def _format_confirmation_response(self, result: ProcessingResult) -> str:
        """Format confirmation required response."""
        lines = []
        
        # Confirmation header
        lines.append(f"{self.colors['prompt']}❓ CONFIRMATION REQUIRED{self.colors['reset']}")
        
        # Main message
        if result.message:
            lines.append(f"{self.colors['bold']}{result.message}{self.colors['reset']}")
        
        # Operation details
        if result.operation_details:
            lines.append("")
            lines.append(f"{self.colors['info']}📋 Operation Details:{self.colors['reset']}")
            
            operation = result.operation_details.get('operation', 'unknown')
            risk_level = result.operation_details.get('risk_level', 'unknown')
            
            lines.append(f"  • Operation: {operation}")
            lines.append(f"  • Risk Level: {self._format_risk_level(risk_level)}")
            
            # Parameters
            parameters = result.operation_details.get('parameters', {})
            if parameters:
                lines.append("  • Parameters:")
                for key, value in parameters.items():
                    if isinstance(value, (dict, list)):
                        value = json.dumps(value, indent=2)
                    lines.append(f"    - {key}: {value}")
        
        # Safety report
        if result.safety_report:
            lines.append("")
            lines.append(f"{self.colors['warning']}🛡️ Safety Analysis:{self.colors['reset']}")
            lines.extend(self._format_safety_report(result.safety_report))
        
        return "\n".join(lines)
    
    def _format_clarification_response(self, result: ProcessingResult) -> str:
        """Format clarification needed response."""
        lines = []
        
        # Clarification header
        lines.append(f"{self.colors['prompt']}❓ CLARIFICATION NEEDED{self.colors['reset']}")
        
        # Main message
        if result.message:
            lines.append(f"{self.colors['bold']}{result.message}{self.colors['reset']}")
        
        # Clarification questions
        if result.clarification_questions:
            lines.append("")
            lines.append(f"{self.colors['info']}Please help me understand:{self.colors['reset']}")
            for i, question in enumerate(result.clarification_questions, 1):
                lines.append(f"  {i}. {question}")
        
        return "\n".join(lines)
    
    def _format_generic_response(self, result: ProcessingResult) -> str:
        """Format generic response for unknown status."""
        lines = []
        
        # Generic header
        status_display = result.status.value.upper().replace('_', ' ')
        lines.append(f"{self.colors['info']}ℹ️ {status_display}{self.colors['reset']}")
        
        # Message
        if result.message:
            lines.append(f"{self.colors['bold']}{result.message}{self.colors['reset']}")
        
        return "\n".join(lines)
    
    def _format_data(self, data: Any) -> List[str]:
        """Format data for display."""
        if data is None:
            return []
        
        lines = []
        
        if isinstance(data, dict):
            lines.extend(self._format_dict_data(data))
        elif isinstance(data, list):
            lines.extend(self._format_list_data(data))
        elif isinstance(data, str):
            lines.append(data)
        else:
            lines.append(str(data))
        
        return lines
    
    def _format_dict_data(self, data: Dict[str, Any]) -> List[str]:
        """Format dictionary data."""
        lines = []
        
        for key, value in data.items():
            if isinstance(value, (dict, list)):
                lines.append(f"{key}:")
                sub_lines = self._format_data(value)
                lines.extend([f"  {line}" for line in sub_lines])
            else:
                lines.append(f"{key}: {value}")
        
        return lines
    
    def _format_list_data(self, data: List[Any]) -> List[str]:
        """Format list data, including tabular data."""
        if not data:
            return ["(no data)"]
        
        lines = []
        
        # Check if this looks like tabular data
        if all(isinstance(item, (list, tuple)) for item in data):
            lines.extend(self._format_table_data(data))
        elif all(isinstance(item, dict) for item in data):
            lines.extend(self._format_records_data(data))
        else:
            # Simple list
            for i, item in enumerate(data[:self.options.max_data_rows], 1):
                lines.append(f"{i:3d}. {item}")
            
            if len(data) > self.options.max_data_rows:
                lines.append(f"... and {len(data) - self.options.max_data_rows} more rows")
        
        return lines
    
    def _format_table_data(self, data: List[List[Any]]) -> List[str]:
        """Format tabular data with columns."""
        if not data:
            return []
        
        lines = []
        
        # Limit rows displayed
        display_data = data[:self.options.max_data_rows]
        
        # Calculate column widths
        if display_data:
            max_cols = max(len(row) for row in display_data)
            col_widths = []
            
            for col in range(max_cols):
                max_width = 0
                for row in display_data:
                    if col < len(row):
                        cell_width = len(str(row[col]))
                        max_width = max(max_width, cell_width)
                col_widths.append(min(max_width, self.options.max_column_width))
            
            # Format rows
            for i, row in enumerate(display_data):
                formatted_cells = []
                for j, cell in enumerate(row):
                    if j < len(col_widths):
                        cell_str = str(cell)[:self.options.max_column_width]
                        formatted_cells.append(cell_str.ljust(col_widths[j]))
                    else:
                        formatted_cells.append(str(cell))
                
                row_str = " | ".join(formatted_cells)
                lines.append(f"{i+1:3d}: {row_str}")
                
                # Add separator after header (first row)
                if i == 0 and len(display_data) > 1:
                    separator = "-+-".join(["-" * width for width in col_widths])
                    lines.append(f"     {separator}")
        
        if len(data) > self.options.max_data_rows:
            lines.append(f"... and {len(data) - self.options.max_data_rows} more rows")
        
        return lines
    
    def _format_records_data(self, data: List[Dict[str, Any]]) -> List[str]:
        """Format list of dictionary records."""
        lines = []
        
        display_data = data[:self.options.max_data_rows]
        
        for i, record in enumerate(display_data, 1):
            lines.append(f"Record {i}:")
            for key, value in record.items():
                lines.append(f"  {key}: {value}")
            if i < len(display_data):
                lines.append("")  # Blank line between records
        
        if len(data) > self.options.max_data_rows:
            lines.append(f"... and {len(data) - self.options.max_data_rows} more records")
        
        return lines
    
    def _format_safety_report(self, safety_report: str) -> List[str]:
        """Format safety report for display."""
        # For now, just split by lines and indent
        lines = safety_report.split('\n')
        return [f"  {line}" for line in lines if line.strip()]
    
    def _format_risk_level(self, risk_level: str) -> str:
        """Format risk level with appropriate coloring."""
        risk_colors = {
            'low': self.colors['success'],
            'medium': self.colors['warning'],
            'high': self.colors['error']
        }
        
        color = risk_colors.get(risk_level.lower(), self.colors['info'])
        return f"{color}{risk_level.upper()}{self.colors['reset']}"
    
    def format_confirmation_prompt(self, result: ProcessingResult) -> str:
        """Format confirmation prompt specifically."""
        if not result.confirmation_prompt:
            return "Do you want to proceed? (yes/no)"
        
        lines = []
        lines.append(f"{self.colors['prompt']}{result.confirmation_prompt}{self.colors['reset']}")
        lines.append(f"{self.colors['dim']}Type 'yes' to proceed or 'no' to cancel{self.colors['reset']}")
        
        return "\n".join(lines)
    
    def format_data_summary(self, data: Any, title: str = "Data Summary") -> str:
        """Format a data summary for quick display."""
        lines = []
        lines.append(f"{self.colors['info']}{title}:{self.colors['reset']}")
        
        if isinstance(data, list):
            lines.append(f"  Items: {len(data)}")
            if data and isinstance(data[0], dict):
                lines.append(f"  Columns: {', '.join(data[0].keys())}")
        elif isinstance(data, dict):
            lines.append(f"  Keys: {', '.join(data.keys())}")
        else:
            lines.append(f"  Type: {type(data).__name__}")
            lines.append(f"  Value: {str(data)[:100]}...")
        
        return "\n".join(lines)


def create_response_formatter(use_colors: bool = True, compact_mode: bool = False) -> ResponseFormatter:
    """
    Factory function to create a response formatter.
    
    Args:
        use_colors: Whether to use terminal colors
        compact_mode: Whether to use compact formatting
        
    Returns:
        ResponseFormatter: Configured response formatter
    """
    options = FormattingOptions(
        use_colors=use_colors,
        compact_mode=compact_mode
    )
    return ResponseFormatter(options)