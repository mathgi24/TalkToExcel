"""
Confirmation handling for operations that require user approval.

This module handles the display and processing of confirmation prompts
for operations that require explicit user consent due to safety considerations.
"""

from typing import Optional, Dict, Any, List
from dataclasses import dataclass
from enum import Enum

from processing.command_processor import ProcessingResult

# Add src directory to Python path for imports
import sys
from pathlib import Path
current_dir = Path(__file__).parent
src_dir = current_dir.parent if current_dir.name != 'src' else current_dir
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))



class ConfirmationResponse(Enum):
    """User's response to confirmation prompt."""
    YES = "yes"
    NO = "no"
    CANCEL = "cancel"
    UNKNOWN = "unknown"


@dataclass
class ConfirmationContext:
    """Context information for confirmation prompts."""
    operation_name: str
    risk_level: str
    affected_items: Optional[int] = None
    affected_sheets: Optional[List[str]] = None
    parameters: Optional[Dict[str, Any]] = None
    safety_warnings: Optional[List[str]] = None


class ConfirmationHandler:
    """
    Handles confirmation prompts and user responses.
    
    Provides clear confirmation prompts for operations that require
    user approval, especially those with safety implications.
    """
    
    def __init__(self):
        """Initialize confirmation handler."""
        # Positive confirmation phrases
        self.positive_responses = {
            'yes', 'y', 'ok', 'okay', 'proceed', 'continue', 'confirm', 
            'go', 'do it', 'execute', 'run', 'sure', 'absolutely',
            'affirmative', 'correct', 'right', 'true', '1'
        }
        
        # Negative confirmation phrases
        self.negative_responses = {
            'no', 'n', 'cancel', 'stop', 'abort', 'quit', 'exit',
            'deny', 'refuse', 'reject', 'negative', 'false', '0',
            'nope', 'nah', 'never', 'skip', 'not'
        }
        
        # Risk level display configuration
        self.risk_colors = {
            'low': '🟢',
            'medium': '🟡', 
            'high': '🔴'
        }
    
    def display_confirmation_prompt(self, result: ProcessingResult) -> None:
        """
        Display a confirmation prompt to the user.
        
        Args:
            result: Processing result containing confirmation details
        """
        print("\n" + "=" * 50)
        print("⚠️  CONFIRMATION REQUIRED")
        print("=" * 50)
        
        # Extract confirmation context
        context = self._extract_confirmation_context(result)
        
        # Display operation details
        self._display_operation_details(context)
        
        # Display safety information
        if context.safety_warnings:
            self._display_safety_warnings(context.safety_warnings)
        
        # Display the confirmation prompt
        if result.confirmation_prompt:
            print(f"\n❓ {result.confirmation_prompt}")
        else:
            print(f"\n❓ Do you want to proceed with this {context.operation_name} operation?")
        
        # Display response instructions
        print("\n💡 Response options:")
        print("   • Type 'yes' or 'y' to proceed")
        print("   • Type 'no' or 'n' to cancel")
        print("   • You can also use phrases like 'ok', 'cancel', 'proceed', etc.")
        
        print("\n" + "-" * 50)
    
    def _extract_confirmation_context(self, result: ProcessingResult) -> ConfirmationContext:
        """Extract confirmation context from processing result."""
        operation_details = result.operation_details or {}
        
        return ConfirmationContext(
            operation_name=operation_details.get('operation', 'unknown operation'),
            risk_level=operation_details.get('risk_level', 'unknown'),
            parameters=operation_details.get('parameters', {}),
            safety_warnings=result.warnings
        )
    
    def _display_operation_details(self, context: ConfirmationContext) -> None:
        """Display operation details for confirmation."""
        print(f"\n🔧 Operation: {context.operation_name}")
        
        # Risk level with visual indicator
        risk_icon = self.risk_colors.get(context.risk_level.lower(), '❓')
        print(f"⚠️  Risk Level: {risk_icon} {context.risk_level.upper()}")
        
        # Parameters (filtered for user-relevant info)
        if context.parameters:
            user_relevant_params = self._filter_user_relevant_parameters(context.parameters)
            if user_relevant_params:
                print("\n📋 Operation Details:")
                for key, value in user_relevant_params.items():
                    print(f"   • {self._format_parameter_name(key)}: {self._format_parameter_value(value)}")
    
    def _filter_user_relevant_parameters(self, parameters: Dict[str, Any]) -> Dict[str, Any]:
        """Filter parameters to show only user-relevant information."""
        # Parameters that are relevant to show users
        relevant_keys = {
            'sheet_name', 'target_sheet', 'data_range', 'conditions',
            'affected_rows', 'affected_columns', 'chart_type', 'title',
            'columns', 'values', 'limit', 'sort_by'
        }
        
        filtered = {}
        for key, value in parameters.items():
            if key in relevant_keys and value is not None:
                filtered[key] = value
        
        return filtered
    
    def _format_parameter_name(self, param_name: str) -> str:
        """Format parameter name for display."""
        # Convert snake_case to readable format
        return param_name.replace('_', ' ').title()
    
    def _format_parameter_value(self, value: Any) -> str:
        """Format parameter value for display."""
        if isinstance(value, list):
            if len(value) <= 3:
                return ', '.join(str(v) for v in value)
            else:
                return f"{', '.join(str(v) for v in value[:3])} (and {len(value) - 3} more)"
        elif isinstance(value, dict):
            return f"{len(value)} items"
        elif isinstance(value, str) and len(value) > 50:
            return f"{value[:47]}..."
        else:
            return str(value)
    
    def _display_safety_warnings(self, warnings: List[str]) -> None:
        """Display safety warnings."""
        print("\n🛡️ Safety Considerations:")
        for warning in warnings:
            print(f"   ⚠️ {warning}")
    
    def parse_confirmation_response(self, user_input: str) -> Optional[bool]:
        """
        Parse user's confirmation response.
        
        Args:
            user_input: User's response string
            
        Returns:
            True for positive response, False for negative, None for unclear
        """
        if not user_input:
            return None
        
        # Normalize input
        normalized = user_input.lower().strip()
        
        # Check for positive responses
        if normalized in self.positive_responses:
            return True
        
        # Check for negative responses
        if normalized in self.negative_responses:
            return False
        
        # Check for partial matches in longer responses
        words = normalized.split()
        
        # Look for positive indicators
        positive_found = any(word in self.positive_responses for word in words)
        negative_found = any(word in self.negative_responses for word in words)
        
        if positive_found and not negative_found:
            return True
        elif negative_found and not positive_found:
            return False
        
        # Ambiguous or unclear response
        return None
    
    def get_confirmation_response_type(self, user_input: str) -> ConfirmationResponse:
        """
        Get the type of confirmation response.
        
        Args:
            user_input: User's response string
            
        Returns:
            ConfirmationResponse enum value
        """
        parsed = self.parse_confirmation_response(user_input)
        
        if parsed is True:
            return ConfirmationResponse.YES
        elif parsed is False:
            return ConfirmationResponse.NO
        else:
            return ConfirmationResponse.UNKNOWN
    
    def generate_confirmation_help(self) -> str:
        """Generate help text for confirmation responses."""
        return """
Confirmation Response Help:

✅ To PROCEED with the operation, you can say:
   • yes, y, ok, okay, proceed, continue, confirm
   • go, do it, execute, run, sure, absolutely

❌ To CANCEL the operation, you can say:
   • no, n, cancel, stop, abort, quit
   • deny, refuse, reject, skip

💡 Tips:
   • Be clear and direct with your response
   • If unsure, type 'no' to cancel safely
   • You can always try the command again later
"""
    
    def create_safety_confirmation_prompt(self, operation: str, risk_level: str, 
                                        details: Optional[Dict[str, Any]] = None) -> str:
        """
        Create a safety-focused confirmation prompt.
        
        Args:
            operation: Name of the operation
            risk_level: Risk level (low, medium, high)
            details: Additional operation details
            
        Returns:
            Formatted confirmation prompt
        """
        risk_icon = self.risk_colors.get(risk_level.lower(), '❓')
        
        prompt_parts = [
            f"This {operation} operation has a {risk_icon} {risk_level.upper()} risk level."
        ]
        
        if risk_level.lower() == 'high':
            prompt_parts.append("This operation could significantly modify your data.")
        elif risk_level.lower() == 'medium':
            prompt_parts.append("This operation will make changes to your data.")
        
        if details:
            if 'affected_rows' in details:
                prompt_parts.append(f"It will affect approximately {details['affected_rows']} rows.")
            if 'sheet_name' in details:
                prompt_parts.append(f"Target sheet: {details['sheet_name']}")
        
        prompt_parts.append("Do you want to proceed?")
        
        return " ".join(prompt_parts)
    
    def validate_confirmation_context(self, result: ProcessingResult) -> bool:
        """
        Validate that a processing result contains proper confirmation context.
        
        Args:
            result: Processing result to validate
            
        Returns:
            True if context is valid for confirmation
        """
        if not result.operation_details:
            return False
        
        required_fields = ['confirmation_id', 'operation']
        return all(field in result.operation_details for field in required_fields)


def create_confirmation_handler() -> ConfirmationHandler:
    """
    Factory function to create a confirmation handler.
    
    Returns:
        ConfirmationHandler: Ready-to-use confirmation handler
    """
    return ConfirmationHandler()