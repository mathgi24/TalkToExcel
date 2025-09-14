"""
Command-line interface for natural language input.

This module provides the main CLI interface for users to interact with the
Excel-LLM Integration Tool using natural language commands.
"""

import sys
import os
from typing import Optional, List, Dict, Any

# Add src directory to Python path for imports
import sys
from pathlib import Path
current_dir = Path(__file__).parent
src_dir = current_dir.parent if current_dir.name != 'src' else current_dir
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))


# Import readline with Windows compatibility
try:
    import readline
except ImportError:
    try:
        import pyreadline3 as readline
    except ImportError:
        try:
            import pyreadline as readline
        except ImportError:
            # If no readline is available, create a dummy module
            class DummyReadline:
                def parse_and_bind(self, *args): pass
                def add_history(self, *args): pass
                def set_completer(self, *args): pass
                def set_completer_delims(self, *args): pass
            readline = DummyReadline()
from pathlib import Path
from dataclasses import dataclass
from enum import Enum

from processing.command_processor import CommandProcessor, ProcessingStatus, ProcessingResult
from .response_formatter import ResponseFormatter
from .confirmation_handler import ConfirmationHandler
from .clarification_handler import ClarificationHandler
from excel.excel_service import ExcelService


class CLIState(Enum):
    """Current state of the CLI interface."""
    READY = "ready"
    WAITING_FOR_FILE = "waiting_for_file"
    WAITING_FOR_CONFIRMATION = "waiting_for_confirmation"
    WAITING_FOR_CLARIFICATION = "waiting_for_clarification"
    PROCESSING = "processing"
    ERROR = "error"


@dataclass
class CLISession:
    """CLI session state management."""
    current_file: Optional[str] = None
    state: CLIState = CLIState.READY
    pending_confirmation_id: Optional[str] = None
    clarification_context: Optional[Dict[str, Any]] = None
    command_history: List[str] = None
    
    def __post_init__(self):
        if self.command_history is None:
            self.command_history = []


class CLIInterface:
    """
    Command-line interface for natural language Excel operations.
    
    Provides an interactive CLI that allows users to:
    - Load Excel files
    - Execute natural language commands
    - Handle confirmations and clarifications
    - View formatted responses
    """
    
    def __init__(self, command_processor: CommandProcessor):
        """
        Initialize CLI interface.
        
        Args:
            command_processor: Command processor for handling user commands
        """
        self.command_processor = command_processor
        self.excel_service = command_processor.excel_service
        self.response_formatter = ResponseFormatter()
        self.confirmation_handler = ConfirmationHandler()
        self.clarification_handler = ClarificationHandler()
        
        self.session = CLISession()
        
        # CLI configuration
        self.prompt = "excel-llm> "
        self.continuation_prompt = "... "
        self.max_history = 100
        self._completion_matches = []
        
        self._setup_readline()
    
    def _setup_readline(self):
        """Setup readline for command history and completion."""
        try:
            # Enable history
            history_file = Path.home() / ".excel_llm_history"
            if history_file.exists():
                readline.read_history_file(str(history_file))
            
            # Set history length
            readline.set_history_length(self.max_history)
            
            # Setup completion
            readline.set_completer(self._command_completer)
            readline.parse_and_bind("tab: complete")
            
            # Save history on exit
            import atexit
            atexit.register(lambda: readline.write_history_file(str(history_file)))
            
        except ImportError:
            # readline not available on all platforms
            pass
    
    def _command_completer(self, text: str, state: int) -> Optional[str]:
        """Provide command completion suggestions."""
        if state == 0:
            # Common command starters
            commands = [
                "load ", "open ", "show ", "display ", "get ", "find ",
                "create ", "add ", "insert ", "new ",
                "update ", "modify ", "change ", "edit ",
                "delete ", "remove ", "drop ",
                "chart ", "plot ", "graph ", "visualize ",
                "help", "exit", "quit", "clear"
            ]
            
            self._completion_matches = [
                cmd for cmd in commands if cmd.startswith(text.lower())
            ]
        
        try:
            return self._completion_matches[state]
        except IndexError:
            return None
    
    def start(self):
        """Start the interactive CLI session."""
        self._print_welcome()
        
        try:
            while True:
                try:
                    # Get user input based on current state
                    user_input = self._get_user_input()
                    
                    if not user_input.strip():
                        continue
                    
                    # Handle special commands
                    if self._handle_special_commands(user_input):
                        continue
                    
                    # Process the command
                    self._process_user_input(user_input)
                    
                except KeyboardInterrupt:
                    print("\n\nUse 'exit' or 'quit' to leave the program.")
                    continue
                except EOFError:
                    print("\nGoodbye!")
                    break
                    
        except Exception as e:
            print(f"\nUnexpected error: {e}")
            sys.exit(1)
    
    def _print_welcome(self):
        """Print welcome message and instructions."""
        print("=" * 60)
        print("🚀 Excel-LLM Integration Tool")
        print("=" * 60)
        print()
        print("Welcome! I can help you work with Excel files using natural language.")
        print()
        print("Getting started:")
        print("  • Load a file: 'load myfile.xlsx' or 'open data.csv'")
        print("  • Ask questions: 'show me the sales data'")
        print("  • Create charts: 'create a bar chart from the revenue column'")
        print("  • Modify data: 'add a new row with product Laptop, price 1200'")
        print()
        print("Commands:")
        print("  • help    - Show detailed help")
        print("  • clear   - Clear screen")
        print("  • exit    - Exit the program")
        print()
        
        if not self.session.current_file:
            print("💡 Start by loading an Excel file to begin working with your data.")
        
        print()
    
    def _get_user_input(self) -> str:
        """Get user input based on current CLI state."""
        try:
            if self.session.state == CLIState.WAITING_FOR_FILE:
                prompt = "📁 Enter file path> "
            elif self.session.state == CLIState.WAITING_FOR_CONFIRMATION:
                prompt = "❓ Confirm (yes/no)> "
            elif self.session.state == CLIState.WAITING_FOR_CLARIFICATION:
                prompt = "❓ Please clarify> "
            else:
                # Show current file in prompt if loaded
                if self.session.current_file:
                    filename = Path(self.session.current_file).name
                    prompt = f"excel-llm [{filename}]> "
                else:
                    prompt = self.prompt
            
            return input(prompt).strip()
            
        except (KeyboardInterrupt, EOFError):
            raise
    
    def _handle_special_commands(self, user_input: str) -> bool:
        """
        Handle special CLI commands.
        
        Returns:
            True if command was handled, False otherwise
        """
        command = user_input.lower().strip()
        
        if command in ['exit', 'quit', 'q']:
            print("Goodbye!")
            sys.exit(0)
        
        elif command in ['help', 'h', '?']:
            self._show_help()
            return True
        
        elif command in ['clear', 'cls']:
            os.system('cls' if os.name == 'nt' else 'clear')
            return True
        
        elif command.startswith('load ') or command.startswith('open '):
            file_path = command.split(' ', 1)[1] if ' ' in command else ''
            self._load_file(file_path)
            return True
        
        elif command == 'status':
            self._show_status()
            return True
        
        elif command == 'operations':
            self._show_available_operations()
            return True
        
        elif command == 'history':
            self._show_command_history()
            return True
        
        return False
    
    def _process_user_input(self, user_input: str):
        """Process user input based on current state."""
        self.session.command_history.append(user_input)
        
        if self.session.state == CLIState.WAITING_FOR_FILE:
            self._load_file(user_input)
            
        elif self.session.state == CLIState.WAITING_FOR_CONFIRMATION:
            self._handle_confirmation_response(user_input)
            
        elif self.session.state == CLIState.WAITING_FOR_CLARIFICATION:
            self._handle_clarification_response(user_input)
            
        else:
            # Regular command processing
            if not self.session.current_file:
                print("⚠️  No Excel file loaded. Use 'load <filename>' to load a file first.")
                return
            
            self._execute_command(user_input)
    
    def _load_file(self, file_path: str):
        """Load an Excel file."""
        if not file_path:
            print("📁 Please specify a file path.")
            self.session.state = CLIState.WAITING_FOR_FILE
            return
        
        try:
            # Expand user path and resolve relative paths
            file_path = str(Path(file_path).expanduser().resolve())
            
            if not Path(file_path).exists():
                print(f"❌ File not found: {file_path}")
                self.session.state = CLIState.READY
                return
            
            print(f"📂 Loading file: {file_path}")
            self.excel_service.load_workbook(file_path)
            
            self.session.current_file = file_path
            self.session.state = CLIState.READY
            
            # Show file information
            structure = self.excel_service.get_structure()
            if structure:
                print(f"✅ File loaded successfully!")
                print(f"   📊 Sheets: {', '.join(structure.sheets)}")
                
                # Show first sheet info
                if structure.sheets:
                    first_sheet = structure.sheets[0]
                    sheet_info = structure.get_sheet_info(first_sheet)
                    if sheet_info:
                        print(f"   📋 '{first_sheet}': {sheet_info.get('row_count', 0)} rows")
                        if 'headers' in sheet_info:
                            headers = sheet_info['headers'][:5]  # Show first 5 headers
                            headers_str = ', '.join(headers)
                            if len(sheet_info['headers']) > 5:
                                headers_str += f" (and {len(sheet_info['headers']) - 5} more)"
                            print(f"   📝 Columns: {headers_str}")
            
            print("\n💡 You can now ask questions about your data!")
            print("   Examples:")
            print("   • 'show me the first 10 rows'")
            print("   • 'create a chart from the sales data'")
            print("   • 'add a new row with...'")
            
        except Exception as e:
            print(f"❌ Error loading file: {e}")
            self.session.state = CLIState.READY
    
    def _execute_command(self, user_command: str):
        """Execute a natural language command."""
        self.session.state = CLIState.PROCESSING
        
        try:
            print("🤔 Processing your request...")
            
            result = self.command_processor.process_command(user_command)
            
            # Format and display the result
            self._display_result(result)
            
            # Update session state based on result
            if result.status == ProcessingStatus.CONFIRMATION_REQUIRED:
                self.session.state = CLIState.WAITING_FOR_CONFIRMATION
                self.session.pending_confirmation_id = result.operation_details.get('confirmation_id')
            elif result.status == ProcessingStatus.CLARIFICATION_NEEDED:
                self.session.state = CLIState.WAITING_FOR_CLARIFICATION
                self.session.clarification_context = {
                    'original_command': user_command,
                    'questions': result.clarification_questions
                }
            else:
                self.session.state = CLIState.READY
                
        except Exception as e:
            print(f"❌ Error processing command: {e}")
            self.session.state = CLIState.READY
    
    def _display_result(self, result: ProcessingResult):
        """Display the processing result to the user."""
        formatted_response = self.response_formatter.format_response(result)
        print(formatted_response)
        
        # Handle specific result types
        if result.status == ProcessingStatus.CONFIRMATION_REQUIRED:
            self.confirmation_handler.display_confirmation_prompt(result)
        elif result.status == ProcessingStatus.CLARIFICATION_NEEDED:
            self.clarification_handler.display_clarification_questions(result)
    
    def _handle_confirmation_response(self, user_input: str):
        """Handle user's response to confirmation prompt."""
        confirmed = self.confirmation_handler.parse_confirmation_response(user_input)
        
        if confirmed is None:
            print("❓ Please answer 'yes' or 'no'")
            return
        
        # Process confirmation response
        result = self.command_processor.process_command(
            user_input,
            confirmation_id=self.session.pending_confirmation_id,
            user_confirmed=confirmed
        )
        
        self._display_result(result)
        
        # Reset state
        self.session.state = CLIState.READY
        self.session.pending_confirmation_id = None
    
    def _handle_clarification_response(self, user_input: str):
        """Handle user's response to clarification questions."""
        if not self.session.clarification_context:
            print("❌ No clarification context available")
            self.session.state = CLIState.READY
            return
        
        # Check if this is a chart field clarification
        original_command = self.session.clarification_context['original_command']
        
        # Try to parse chart field selection from user response
        if self._is_chart_field_clarification(original_command, user_input):
            result = self._handle_chart_field_selection(original_command, user_input)
        else:
            # Default behavior: combine original command with clarification
            enhanced_command = f"{original_command}. {user_input}"
            print(f"🔄 Processing enhanced command: {enhanced_command}")
            result = self.command_processor.process_command(enhanced_command)
        
        self._display_result(result)
        
        # Reset state
        self.session.state = CLIState.READY
        self.session.clarification_context = None
    
    def _is_chart_field_clarification(self, original_command: str, user_response: str) -> bool:
        """Check if this is a chart field clarification response."""
        return ("chart" in original_command.lower() and 
                ("use" in user_response.lower() or 
                 "for categories" in user_response.lower() or
                 "for values" in user_response.lower() or
                 "and" in user_response.lower()))
    
    def _handle_chart_field_selection(self, original_command: str, user_response: str) -> 'ProcessingResult':
        """Handle chart field selection from user clarification."""
        # Parse the user's field selection
        # Expected formats:
        # "Use Item for categories and Stock for values"
        # "Item and Stock"
        # "categories: Item, values: Stock"
        
        category_field = None
        value_field = None
        
        # Try different parsing patterns
        user_lower = user_response.lower()
        
        # Pattern 1: "use X for categories and Y for values"
        if "use" in user_lower and "for categories" in user_lower and "for values" in user_lower:
            parts = user_lower.split("for categories")
            if len(parts) >= 2:
                category_part = parts[0].replace("use", "").strip()
                value_part = parts[1].split("for values")[0].replace("and", "").strip()
                category_field = category_part
                value_field = value_part
        
        # Pattern 2: "X and Y" (assume first is category, second is value)
        elif " and " in user_lower:
            parts = user_response.split(" and ")
            if len(parts) >= 2:
                category_field = parts[0].strip()
                value_field = parts[1].strip()
        
        # Pattern 3: "categories: X, values: Y"
        elif "categories:" in user_lower and "values:" in user_lower:
            for part in user_response.split(","):
                part = part.strip()
                if part.lower().startswith("categories:"):
                    category_field = part.split(":", 1)[1].strip()
                elif part.lower().startswith("values:"):
                    value_field = part.split(":", 1)[1].strip()
        
        if category_field and value_field:
            # Create a specific chart command with the selected fields
            # Extract sheet name from original command
            sheet_name = None
            if "inventory" in original_command.lower():
                sheet_name = "Inventory"
            elif "sales" in original_command.lower():
                sheet_name = "Sales Data"
            elif "employee" in original_command.lower():
                sheet_name = "Employees"
            
            if sheet_name:
                # Create specific data range based on field selection
                enhanced_command = f"create pie chart from {sheet_name} using {category_field} and {value_field}"
                print(f"🔄 Creating chart with: categories='{category_field}', values='{value_field}'")
                return self.command_processor.process_command_with_fields(
                    original_command, sheet_name, category_field, value_field
                )
        
        # Fallback to default behavior
        enhanced_command = f"{original_command}. {user_response}"
        print(f"🔄 Processing enhanced command: {enhanced_command}")
        return self.command_processor.process_command(enhanced_command)
    
    def _show_help(self):
        """Show detailed help information."""
        print("\n" + "=" * 60)
        print("📖 HELP - Excel-LLM Integration Tool")
        print("=" * 60)
        print()
        print("🎯 NATURAL LANGUAGE COMMANDS:")
        print()
        print("📊 Data Operations:")
        print("  • 'show me the sales data'")
        print("  • 'find all products with price > 100'")
        print("  • 'get the top 10 customers by revenue'")
        print("  • 'add a new row: Product=Laptop, Price=1200, Category=Electronics'")
        print("  • 'update the price of Product X to 150'")
        print("  • 'delete all rows where status is inactive'")
        print()
        print("📈 Visualization:")
        print("  • 'create a bar chart from the sales column'")
        print("  • 'make a pie chart showing revenue by region'")
        print("  • 'plot a line graph of monthly trends'")
        print("  • 'shift the chart left by 2 units'")
        print()
        print("⚙️ SYSTEM COMMANDS:")
        print("  • load <file>     - Load an Excel file")
        print("  • status          - Show current file and system status")
        print("  • operations      - List available operations")
        print("  • history         - Show command history")
        print("  • clear           - Clear screen")
        print("  • help            - Show this help")
        print("  • exit            - Exit the program")
        print()
        print("💡 TIPS:")
        print("  • Be specific about which data you want to work with")
        print("  • Mention sheet names if your file has multiple sheets")
        print("  • The system will ask for clarification if needed")
        print("  • Dangerous operations require confirmation")
        print()
    
    def _show_status(self):
        """Show current system status."""
        print("\n📊 SYSTEM STATUS")
        print("-" * 30)
        
        # File status
        if self.session.current_file:
            print(f"📁 Current file: {Path(self.session.current_file).name}")
            print(f"   Full path: {self.session.current_file}")
            
            structure = self.excel_service.get_structure()
            if structure:
                print(f"   📋 Sheets: {len(structure.sheets)}")
                for sheet in structure.sheets:
                    sheet_info = structure.get_sheet_info(sheet)
                    row_count = sheet_info.get('row_count', 0) if sheet_info else 0
                    print(f"      • {sheet}: {row_count} rows")
        else:
            print("📁 No file loaded")
        
        # Session status
        print(f"🔄 State: {self.session.state.value}")
        print(f"📝 Commands in history: {len(self.session.command_history)}")
        
        # Available operations
        operations = self.command_processor.get_available_operations()
        total_ops = sum(len(ops) for ops in operations.values())
        print(f"⚙️ Available operations: {total_ops}")
        
        print()
    
    def _show_available_operations(self):
        """Show available operations by category."""
        print("\n⚙️ AVAILABLE OPERATIONS")
        print("-" * 40)
        
        operations = self.command_processor.get_available_operations()
        
        for category, ops in operations.items():
            category_name = category.replace('_', ' ').title()
            print(f"\n📂 {category_name}:")
            
            for op in ops:
                # Get operation help if available
                help_info = self.command_processor.get_operation_help(category, op)
                if help_info:
                    safety_level = help_info.get('safety_level', 'unknown')
                    safety_icon = {'safe': '✅', 'medium': '⚠️', 'high': '🚫'}.get(safety_level, '❓')
                    print(f"  {safety_icon} {op}")
                    
                    examples = help_info.get('examples', [])
                    if examples:
                        print(f"      Example: {examples[0]}")
                else:
                    print(f"  • {op}")
        
        print()
    
    def _show_command_history(self):
        """Show recent command history."""
        print("\n📝 COMMAND HISTORY")
        print("-" * 30)
        
        if not self.session.command_history:
            print("No commands in history")
            return
        
        # Show last 10 commands
        recent_commands = self.session.command_history[-10:]
        
        for i, command in enumerate(recent_commands, 1):
            print(f"{i:2d}. {command}")
        
        if len(self.session.command_history) > 10:
            print(f"... and {len(self.session.command_history) - 10} more")
        
        print()


def create_cli_interface(command_processor: CommandProcessor) -> CLIInterface:
    """
    Factory function to create a CLI interface.
    
    Args:
        command_processor: Initialized command processor
        
    Returns:
        CLIInterface: Ready-to-use CLI interface
    """
    return CLIInterface(command_processor)