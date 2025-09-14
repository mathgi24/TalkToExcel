"""
Main Excel-LLM Integration System

This module provides the main system class that integrates all components
for end-to-end Excel operations using natural language commands.

Requirements: 7.1, 7.2, 7.3
"""

import os
import sys
import logging
from typing import Dict, Any, Optional
from dataclasses import dataclass
from pathlib import Path

# Add src directory to Python path for imports
current_dir = Path(__file__).parent
src_dir = current_dir
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

from config.config_manager import ConfigManager
from llm.ollama_service import OllamaService
from excel.excel_service import ExcelService
from processing.command_processor import CommandProcessor, ProcessingStatus
from ui.cli_interface import CLIInterface
from safety.safety_manager import SafetyManager
from templates.template_registry import TemplateRegistry
from processing.error_handler import ErrorHandler


@dataclass
class OperationResult:
    """Result of an Excel operation"""
    success: bool
    message: str
    data: Optional[Any] = None
    chart_reference: Optional[str] = None
    affected_rows: int = 0
    operation_type: str = ""


class ExcelLLMSystem:
    """Main system class integrating all components"""
    
    def __init__(self, config: Optional[Dict[str, Any]] = None):
        """Initialize the Excel-LLM integration system"""
        self.logger = logging.getLogger(__name__)
        
        # Load configuration
        if config:
            self.config = config
        else:
            config_manager = ConfigManager()
            # Build config dictionary from ConfigManager methods
            self.config = {
                'ollama': config_manager.get_ollama_config(),
                'backup': config_manager.get_backup_config(),
                'safety': config_manager.get_safety_config(),
                'excel': config_manager.get_excel_config(),
                'logging': config_manager.get_logging_config()
            }
        
        # Initialize core components
        self._initialize_components()
        
        # System status
        self.is_initialized = False
        self._check_system_health()
    
    def _initialize_components(self):
        """Initialize all system components"""
        try:
            # Template system (initialize first as other services depend on it)
            self.template_registry = TemplateRegistry()
            
            # Core services
            self.llm_service = OllamaService(template_registry=self.template_registry)
            self.excel_service = ExcelService(
                backup_dir=self.config.get('backup', {}).get('directory', './backups'),
                max_backups=self.config.get('backup', {}).get('retention', 10)
            )
            self.safety_manager = SafetyManager(
                max_rows=self.config.get('safety', {}).get('max_rows', 50),
                max_columns=self.config.get('safety', {}).get('max_columns', 20)
            )
            
            # Processing components
            self.command_processor = CommandProcessor(
                llm_service=self.llm_service,
                template_registry=self.template_registry,
                safety_manager=self.safety_manager,
                excel_service=self.excel_service
            )
            
            # Error handling
            self.error_handler = ErrorHandler()
            
            # User interface
            self.cli_interface = CLIInterface(self.command_processor)
            
            self.logger.info("All system components initialized successfully")
            
        except Exception as e:
            self.logger.error(f"Failed to initialize system components: {e}")
            raise
    
    def _check_system_health(self):
        """Check if all system components are healthy"""
        try:
            # Check LLM service
            if not self.llm_service.is_available():
                self.logger.warning("LLM service is not available")
                return False
            
            # Check template registry
            if not self.template_registry.is_loaded():
                self.logger.warning("Template registry is not loaded")
                return False
            
            self.is_initialized = True
            self.logger.info("System health check passed")
            return True
            
        except Exception as e:
            self.logger.error(f"System health check failed: {e}")
            return False
    
    def process_command(self, command: str, file_path: str) -> OperationResult:
        """
        Process a natural language command on an Excel file
        
        Args:
            command: Natural language command
            file_path: Path to Excel file
            
        Returns:
            OperationResult with success status and details
        """
        if not self.is_initialized:
            return OperationResult(
                success=False,
                message="System is not properly initialized"
            )
        
        try:
            # Validate file exists
            if not os.path.exists(file_path):
                return OperationResult(
                    success=False,
                    message=f"File not found: {file_path}"
                )
            
            # Process the command
            result = self.command_processor.process_command(command, file_path)
            
            # Convert ProcessingResult to OperationResult format
            success = result.status == ProcessingStatus.SUCCESS
            return OperationResult(
                success=success,
                message=result.message,
                data=result.data,
                chart_reference=getattr(result, 'chart_reference', None),
                affected_rows=getattr(result, 'affected_rows', 0),
                operation_type=getattr(result.operation_details, 'operation_type', 'unknown') if result.operation_details else 'unknown'
            )
            
        except Exception as e:
            self.logger.error(f"Error processing command '{command}': {e}")
            error_info = self.error_handler.handle_error(e, {
                'command': command,
                'file_path': file_path
            })
            return OperationResult(
                success=False,
                message=error_info.user_message,
                operation_type='error'
            )
    
    def start_interactive_session(self, file_path: Optional[str] = None):
        """Start an interactive CLI session"""
        if not self.is_initialized:
            print("System is not properly initialized. Please check configuration.")
            return
        
        try:
            # Load file if provided
            if file_path:
                if os.path.exists(file_path):
                    self.excel_service.load_workbook(file_path)
                    print(f"Loaded file: {file_path}")
                else:
                    print(f"Warning: File not found: {file_path}")
            
            self.cli_interface.start()
        except KeyboardInterrupt:
            print("\nSession ended by user")
        except Exception as e:
            self.logger.error(f"Error in interactive session: {e}")
            print(f"Session error: {e}")
    
    def get_system_status(self) -> Dict[str, Any]:
        """Get current system status"""
        return {
            'initialized': self.is_initialized,
            'llm_available': self.llm_service.is_available() if hasattr(self, 'llm_service') else False,
            'templates_loaded': self.template_registry.is_loaded() if hasattr(self, 'template_registry') else False,
            'config_valid': bool(self.config),
            'components': {
                'llm_service': hasattr(self, 'llm_service'),
                'excel_service': hasattr(self, 'excel_service'),
                'command_processor': hasattr(self, 'command_processor'),
                'safety_manager': hasattr(self, 'safety_manager'),
                'template_registry': hasattr(self, 'template_registry'),
                'cli_interface': hasattr(self, 'cli_interface')
            }
        }
    
    def restore_from_backup(self, file_path: str, backup_path: str) -> OperationResult:
        """Restore a file from backup"""
        try:
            success = self.excel_service.restore_from_backup(file_path, backup_path)
            
            if success:
                return OperationResult(
                    success=True,
                    message=f"Successfully restored {file_path} from backup {backup_path}"
                )
            else:
                return OperationResult(
                    success=False,
                    message=f"Failed to restore {file_path} from backup"
                )
                
        except Exception as e:
            self.logger.error(f"Error restoring from backup: {e}")
            return OperationResult(
                success=False,
                message=f"Error during backup restoration: {e}"
            )
    
    def list_available_operations(self) -> Dict[str, Any]:
        """List all available operations"""
        if not hasattr(self, 'template_registry'):
            return {}
        
        return self.template_registry.get_all_operations()
    
    def reload_templates(self):
        """Reload operation templates from YAML files."""
        try:
            self.template_registry.reload_registry()
            self.llm_service._load_operation_templates()
            self.logger.info("Templates reloaded successfully")
            return True
        except Exception as e:
            self.logger.error(f"Failed to reload templates: {e}")
            return False
    
    def shutdown(self):
        """Gracefully shutdown the system"""
        self.logger.info("Shutting down Excel-LLM system")
        
        try:
            # Cleanup components
            if hasattr(self, 'llm_service'):
                self.llm_service.cleanup()
            
            if hasattr(self, 'excel_service'):
                self.excel_service.cleanup()
            
            if hasattr(self, 'template_registry'):
                self.template_registry.cleanup()
            
            self.is_initialized = False
            self.logger.info("System shutdown completed")
            
        except Exception as e:
            self.logger.error(f"Error during shutdown: {e}")


def main():
    """Main entry point for command-line usage"""
    import argparse
    
    parser = argparse.ArgumentParser(description="Excel-LLM Integration System")
    parser.add_argument('--file', '-f', help='Excel file to work with')
    parser.add_argument('--command', '-c', help='Single command to execute')
    parser.add_argument('--config', help='Configuration file path')
    parser.add_argument('--interactive', '-i', action='store_true', help='Start interactive session')
    parser.add_argument('--status', action='store_true', help='Show system status')
    
    args = parser.parse_args()
    
    # Initialize system
    try:
        config = None
        if args.config:
            config_manager = ConfigManager(args.config)
            # Build config dictionary from ConfigManager methods
            config = {
                'ollama': config_manager.get_ollama_config(),
                'backup': config_manager.get_backup_config(),
                'safety': config_manager.get_safety_config(),
                'excel': config_manager.get_excel_config(),
                'logging': config_manager.get_logging_config()
            }
        
        system = ExcelLLMSystem(config)
        
        if args.status:
            # Show system status
            status = system.get_system_status()
            print("System Status:")
            print(f"  Initialized: {status['initialized']}")
            print(f"  LLM Available: {status['llm_available']}")
            print(f"  Templates Loaded: {status['templates_loaded']}")
            print(f"  Config Valid: {status['config_valid']}")
            return
        
        if args.command and args.file:
            # Execute single command
            result = system.process_command(args.command, args.file)
            print(f"Result: {result.message}")
            if result.data:
                print(f"Data: {result.data}")
        
        elif args.interactive or args.file:
            # Start interactive session
            system.start_interactive_session(args.file)
        
        else:
            # Show help
            parser.print_help()
    
    except Exception as e:
        print(f"Error: {e}")
        return 1
    
    finally:
        if 'system' in locals():
            system.shutdown()
    
    return 0


if __name__ == "__main__":
    exit(main())
