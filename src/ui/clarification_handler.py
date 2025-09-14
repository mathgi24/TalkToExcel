"""
Clarification handling for ambiguous or unclear user commands.

This module handles the display and processing of clarification requests
when the system needs more information to properly execute user commands.
"""

from typing import List, Dict, Any, Optional, Tuple
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



class ClarificationType(Enum):
    """Types of clarification needed."""
    AMBIGUOUS_TARGET = "ambiguous_target"
    MISSING_PARAMETERS = "missing_parameters"
    UNCLEAR_INTENT = "unclear_intent"
    MULTIPLE_OPTIONS = "multiple_options"
    INVALID_REFERENCE = "invalid_reference"


@dataclass
class ClarificationContext:
    """Context for clarification requests."""
    clarification_type: ClarificationType
    original_command: str
    questions: List[str]
    suggestions: Optional[List[str]] = None
    available_options: Optional[Dict[str, List[str]]] = None
    error_details: Optional[str] = None


class ClarificationHandler:
    """
    Handles clarification requests and user responses.
    
    Provides clear questions and suggestions when user commands
    are ambiguous or need additional information.
    """
    
    def __init__(self):
        """Initialize clarification handler."""
        # Common clarification patterns
        self.clarification_patterns = {
            'sheet_selection': [
                "Which sheet would you like to work with?",
                "Please specify the sheet name.",
                "Your file has multiple sheets - which one should I use?"
            ],
            'column_selection': [
                "Which column(s) are you referring to?",
                "Please specify the column name(s).",
                "Could you clarify which data columns to use?"
            ],
            'data_range': [
                "What data range should I work with?",
                "Please specify which rows or range of data.",
                "Could you be more specific about the data location?"
            ],
            'operation_type': [
                "What specific operation would you like me to perform?",
                "Could you clarify what action you want to take?",
                "Please be more specific about what you'd like to do."
            ],
            'chart_type': [
                "What type of chart would you like?",
                "Please specify the chart type (bar, line, pie, etc.).",
                "Which visualization would work best for your data?"
            ]
        }
    
    def display_clarification_questions(self, result: ProcessingResult) -> None:
        """
        Display clarification questions to the user.
        
        Args:
            result: Processing result containing clarification details
        """
        print("\n" + "=" * 50)
        print("❓ CLARIFICATION NEEDED")
        print("=" * 50)
        
        # Extract clarification context
        context = self._extract_clarification_context(result)
        
        # Display the main message
        if result.message:
            print(f"\n💭 {result.message}")
        
        # Display clarification questions
        if result.clarification_questions:
            print("\n🤔 To help me understand better, please answer:")
            for i, question in enumerate(result.clarification_questions, 1):
                print(f"   {i}. {question}")
        
        # Display suggestions if available
        if context.suggestions:
            self._display_suggestions(context.suggestions)
        
        # Display available options if any
        if context.available_options:
            self._display_available_options(context.available_options)
        
        # Display response instructions
        print("\n💡 How to respond:")
        print("   • Answer the questions with specific details")
        print("   • You can rephrase your original request with more information")
        print("   • Use specific names for sheets, columns, or data ranges")
        
        print("\n" + "-" * 50)
    
    def _extract_clarification_context(self, result: ProcessingResult) -> ClarificationContext:
        """Extract clarification context from processing result."""
        # Determine clarification type based on questions
        clarification_type = self._determine_clarification_type(result.clarification_questions or [])
        
        return ClarificationContext(
            clarification_type=clarification_type,
            original_command="",  # This would be set by the caller
            questions=result.clarification_questions or [],
            suggestions=result.warnings,  # Using warnings field for suggestions
            error_details=result.message
        )
    
    def _determine_clarification_type(self, questions: List[str]) -> ClarificationType:
        """Determine the type of clarification needed based on questions."""
        questions_text = " ".join(questions).lower()
        
        if any(word in questions_text for word in ['sheet', 'worksheet']):
            return ClarificationType.AMBIGUOUS_TARGET
        elif any(word in questions_text for word in ['column', 'field', 'data']):
            return ClarificationType.MISSING_PARAMETERS
        elif any(word in questions_text for word in ['what', 'which', 'how']):
            return ClarificationType.UNCLEAR_INTENT
        elif any(word in questions_text for word in ['multiple', 'several', 'options']):
            return ClarificationType.MULTIPLE_OPTIONS
        else:
            return ClarificationType.UNCLEAR_INTENT
    
    def _display_suggestions(self, suggestions: List[str]) -> None:
        """Display suggestions to help user clarify their request."""
        print("\n💡 Suggestions:")
        for suggestion in suggestions:
            print(f"   • {suggestion}")
    
    def _display_available_options(self, options: Dict[str, List[str]]) -> None:
        """Display available options for user selection."""
        print("\n📋 Available options:")
        
        for category, items in options.items():
            category_name = category.replace('_', ' ').title()
            print(f"\n   {category_name}:")
            
            for item in items[:10]:  # Limit to first 10 items
                print(f"     • {item}")
            
            if len(items) > 10:
                print(f"     ... and {len(items) - 10} more")
    
    def generate_clarification_questions(self, command: str, context: Dict[str, Any]) -> List[str]:
        """
        Generate appropriate clarification questions based on command and context.
        
        Args:
            command: Original user command
            context: Context information (available sheets, columns, etc.)
            
        Returns:
            List of clarification questions
        """
        questions = []
        command_lower = command.lower()
        
        # Check for missing sheet specification
        if context.get('multiple_sheets') and not self._has_sheet_reference(command):
            available_sheets = context.get('sheet_names', [])
            if len(available_sheets) > 1:
                questions.append(f"Which sheet should I work with? Available: {', '.join(available_sheets)}")
        
        # Check for vague data references
        if any(word in command_lower for word in ['data', 'information', 'stuff', 'things']):
            questions.append("Could you be more specific about which data you're referring to?")
            questions.append("Please mention specific column names or data ranges.")
        
        # Check for unclear operations
        if any(word in command_lower for word in ['do something', 'work with', 'handle']):
            questions.append("What specific operation would you like me to perform?")
            questions.append("For example: create, update, delete, show, or visualize?")
        
        # Check for chart requests without type
        if 'chart' in command_lower or 'plot' in command_lower or 'graph' in command_lower:
            if not any(chart_type in command_lower for chart_type in ['bar', 'line', 'pie', 'scatter', 'area']):
                questions.append("What type of chart would you like? (bar, line, pie, scatter, area, etc.)")
        
        # Check for missing column specifications
        if any(word in command_lower for word in ['show', 'display', 'get']) and 'column' not in command_lower:
            available_columns = context.get('column_names', [])
            if available_columns:
                questions.append(f"Which columns should I include? Available: {', '.join(available_columns[:5])}")
        
        # Default questions if none generated
        if not questions:
            questions.extend([
                "Could you provide more details about what you'd like to do?",
                "Please be more specific about the data or operation you're interested in."
            ])
        
        return questions
    
    def _has_sheet_reference(self, command: str) -> bool:
        """Check if command contains a sheet reference."""
        sheet_indicators = ['sheet', 'worksheet', 'tab', 'page']
        command_lower = command.lower()
        return any(indicator in command_lower for indicator in sheet_indicators)
    
    def generate_contextual_suggestions(self, command: str, context: Dict[str, Any]) -> List[str]:
        """
        Generate contextual suggestions based on command and available data.
        
        Args:
            command: Original user command
            context: Context information
            
        Returns:
            List of suggestions
        """
        suggestions = []
        command_lower = command.lower()
        
        # Suggest specific sheets
        if context.get('sheet_names'):
            suggestions.append(f"Try specifying a sheet: 'from {context['sheet_names'][0]} sheet'")
        
        # Suggest specific columns
        if context.get('column_names'):
            columns = context['column_names'][:3]
            suggestions.append(f"Try mentioning specific columns: {', '.join(columns)}")
        
        # Suggest operation types based on command
        if any(word in command_lower for word in ['show', 'see', 'display']):
            suggestions.append("Try: 'show me the first 10 rows from [sheet name]'")
            suggestions.append("Try: 'display all data where [column] equals [value]'")
        
        elif any(word in command_lower for word in ['chart', 'plot', 'graph']):
            suggestions.append("Try: 'create a bar chart from [column name]'")
            suggestions.append("Try: 'make a pie chart showing [data description]'")
        
        elif any(word in command_lower for word in ['add', 'create', 'insert']):
            suggestions.append("Try: 'add a new row with [column1]=value1, [column2]=value2'")
            suggestions.append("Try: 'insert data into [sheet name]: [values]'")
        
        return suggestions
    
    def create_enhanced_command(self, original_command: str, clarification_response: str) -> str:
        """
        Create an enhanced command by combining original command with clarification.
        
        Args:
            original_command: Original user command
            clarification_response: User's clarification response
            
        Returns:
            Enhanced command string
        """
        # Simple combination for now - could be made more sophisticated
        enhanced = f"{original_command}. {clarification_response}"
        
        # Clean up the enhanced command
        enhanced = enhanced.replace('..', '.').strip()
        
        return enhanced
    
    def validate_clarification_response(self, response: str, context: ClarificationContext) -> Tuple[bool, Optional[str]]:
        """
        Validate a clarification response.
        
        Args:
            response: User's clarification response
            context: Clarification context
            
        Returns:
            Tuple of (is_valid, error_message)
        """
        if not response or not response.strip():
            return False, "Please provide a response to help me understand your request."
        
        response_lower = response.lower().strip()
        
        # Check for non-helpful responses
        unhelpful_responses = {
            'i don\'t know', 'not sure', 'whatever', 'anything', 'doesn\'t matter',
            'just do it', 'figure it out', 'you decide'
        }
        
        if response_lower in unhelpful_responses:
            return False, "Please provide specific information to help me process your request."
        
        # Context-specific validation
        if context.clarification_type == ClarificationType.AMBIGUOUS_TARGET:
            if len(response.split()) < 2:
                return False, "Please provide more specific details about what you want to work with."
        
        return True, None
    
    def get_clarification_help(self, clarification_type: ClarificationType) -> str:
        """
        Get help text for specific clarification types.
        
        Args:
            clarification_type: Type of clarification needed
            
        Returns:
            Help text string
        """
        help_texts = {
            ClarificationType.AMBIGUOUS_TARGET: """
Help: Specifying Data Targets

When I ask about targets, please be specific:
• Sheet names: "from the Sales sheet" or "in worksheet Data"
• Column names: "the Revenue column" or "columns A, B, and C"
• Data ranges: "rows 1-10" or "the first 50 rows"

Examples:
• "Show me data from the Sales sheet"
• "Create a chart from the Revenue column in Q1 Data"
• "Update rows where Status equals 'Active'"
""",
            
            ClarificationType.MISSING_PARAMETERS: """
Help: Providing Missing Information

Please include specific details:
• What data to work with
• Which columns or ranges
• What values to use
• Any conditions or filters

Examples:
• "Add a new row: Name=John, Age=30, Department=Sales"
• "Show all records where Revenue > 1000"
• "Create a bar chart from the Monthly Sales column"
""",
            
            ClarificationType.UNCLEAR_INTENT: """
Help: Clarifying Your Intent

Please specify what you want to do:
• View/Display: "show", "display", "get", "find"
• Create/Add: "create", "add", "insert", "new"
• Modify: "update", "change", "modify", "edit"
• Remove: "delete", "remove", "drop"
• Visualize: "chart", "plot", "graph", "visualize"

Examples:
• "Show me the sales data"
• "Create a new customer record"
• "Update the price for Product X"
• "Delete inactive users"
"""
        }
        
        return help_texts.get(clarification_type, "Please provide more specific information about your request.")


def create_clarification_handler() -> ClarificationHandler:
    """
    Factory function to create a clarification handler.
    
    Returns:
        ClarificationHandler: Ready-to-use clarification handler
    """
    return ClarificationHandler()