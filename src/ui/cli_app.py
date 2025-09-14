"""
Main CLI application for Excel-LLM Integration Tool.

This module provides the main entry point for the command-line interface,
integrating all components for a complete user experience.
"""

import sys
import logging
from pathlib import Path
from typing import Optional

# Add src directory to Python path for imports
import sys
from pathlib import Path
current_dir = Path(__file__).parent
src_dir = current_dir.parent if current_dir.name != 'src' else current_dir
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))


# Add src directory to Python path
src_path = Path(__file__).parent.parent
sys.path.insert(0, str(src_path))

from config.config_manager import config
from llm.ollama_service import OllamaService, OllamaConnectionError
from templates.template_loader import TemplateLoader
from templates.template_registry import TemplateRegistry
from templates.prompt_generator import PromptGenerator
from templates.hot_reload import HotReloadManager
from processing.command_processor import CommandProcessor
from safety.safety_manager import SafetyManager
from excel.excel_service import ExcelService
from ..ui.cli_interface import CLIInterface


class CLIApp:
    """
    Main CLI application for Excel-LLM Integration Tool.
    
    Coordinates initialization of all system components and provides
    the main entry point for the interactive CLI experience.
    """
    
    def __init__(self):
        """Initialize the CLI application."""
        self.logger = self._setup_logging()
        
        # System components
        self.llm_service: Optional[OllamaService] = None
        self.template_registry: Optional[TemplateRegistry] = None
        self.command_processor: Optional[CommandProcessor] = None
        self.cli_interface: Optional[CLIInterface] = None
        
        # Initialization status
        self.initialized = False
        self.initialization_errors = []
    
    def _setup_logging(self) -> logging.Logger:
        """Setup logging configuration."""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('logs/cli_app.log'),
                logging.StreamHandler(sys.stdout)
            ]
        )
        return logging.getLogger(__name__)
    
    def initialize(self) -> bool:
        """
        Initialize all system components.
        
        Returns:
            True if initialization successful, False otherwise
        """
        try:
            self.logger.info("Starting Excel-LLM Integration Tool initialization...")
            
            # Step 1: Validate configuration
            if not self._initialize_configuration():
                return False
            
            # Step 2: Initialize template system
            if not self._initialize_template_system():
                return False
            
            # Step 3: Initialize LLM service
            if not self._initialize_llm_service():
                return False
            
            # Step 4: Initialize Excel service
            if not self._initialize_excel_service():
                return False
            
            # Step 5: Initialize safety manager
            if not self._initialize_safety_manager():
                return False
            
            # Step 6: Initialize command processor
            if not self._initialize_command_processor():
                return False
            
            # Step 7: Initialize CLI interface
            if not self._initialize_cli_interface():
                return False
            
            self.initialized = True
            self.logger.info("✅ System initialization completed successfully!")
            return True
            
        except Exception as e:
            self.logger.error(f"❌ Initialization failed: {str(e)}")
            self.initialization_errors.append(str(e))
            return False
    
    def _initialize_configuration(self) -> bool:
        """Initialize and validate configuration."""
        try:
            config.validate_config()
            self.logger.info("✓ Configuration validated")
            return True
        except Exception as e:
            self.logger.error(f"Configuration error: {str(e)}")
            self.initialization_errors.append(f"Configuration: {str(e)}")
            return False
    
    def _initialize_template_system(self) -> bool:
        """Initialize the template system."""
        try:
            template_loader = TemplateLoader()
            self.template_registry = TemplateRegistry(template_loader)
            
            # Initialize prompt generator and hot reload
            prompt_generator = PromptGenerator(template_loader, self.template_registry)
            hot_reload_manager = HotReloadManager(template_loader, self.template_registry, prompt_generator)
            
            registry_stats = self.template_registry.get_registry_stats()
            self.logger.info(f"✓ Template system: {registry_stats['total_operations']} operations loaded")
            return True
            
        except Exception as e:
            self.logger.error(f"Template system error: {str(e)}")
            self.initialization_errors.append(f"Template system: {str(e)}")
            return False
    
    def _initialize_llm_service(self) -> bool:
        """Initialize the LLM service."""
        try:
            self.llm_service = OllamaService()
            
            # Test connection (non-blocking)
            try:
                if self.llm_service.initialize_connection():
                    self.logger.info("✓ LLM service connected to Ollama")
                else:
                    self.logger.warning("⚠️ LLM service initialized but Ollama not connected")
            except OllamaConnectionError:
                self.logger.warning("⚠️ LLM service initialized but Ollama connection failed")
            
            return True
            
        except Exception as e:
            self.logger.error(f"LLM service error: {str(e)}")
            self.initialization_errors.append(f"LLM service: {str(e)}")
            return False
    
    def _initialize_excel_service(self) -> bool:
        """Initialize the Excel service."""
        try:
            self.excel_service = ExcelService()
            self.logger.info("✓ Excel service initialized")
            return True
            
        except Exception as e:
            self.logger.error(f"Excel service error: {str(e)}")
            self.initialization_errors.append(f"Excel service: {str(e)}")
            return False
    
    def _initialize_safety_manager(self) -> bool:
        """Initialize the safety manager."""
        try:
            self.safety_manager = SafetyManager(max_rows=50, max_columns=20)
            self.logger.info("✓ Safety manager initialized")
            return True
            
        except Exception as e:
            self.logger.error(f"Safety manager error: {str(e)}")
            self.initialization_errors.append(f"Safety manager: {str(e)}")
            return False
    
    def _initialize_command_processor(self) -> bool:
        """Initialize the command processor."""
        try:
            self.command_processor = CommandProcessor(
                llm_service=self.llm_service,
                template_registry=self.template_registry,
                safety_manager=self.safety_manager,
                excel_service=self.excel_service
            )
            self.logger.info("✓ Command processor initialized")
            return True
            
        except Exception as e:
            self.logger.error(f"Command processor error: {str(e)}")
            self.initialization_errors.append(f"Command processor: {str(e)}")
            return False
    
    def _initialize_cli_interface(self) -> bool:
        """Initialize the CLI interface."""
        try:
            self.cli_interface = CLIInterface(self.command_processor)
            self.logger.info("✓ CLI interface initialized")
            return True
            
        except Exception as e:
            self.logger.error(f"CLI interface error: {str(e)}")
            self.initialization_errors.append(f"CLI interface: {str(e)}")
            return False
    
    def run(self) -> None:
        """Run the CLI application."""
        if not self.initialized:
            print("❌ System not properly initialized. Cannot start CLI.")
            if self.initialization_errors:
                print("\nInitialization errors:")
                for error in self.initialization_errors:
                    print(f"  • {error}")
            sys.exit(1)
        
        try:
            self.logger.info("Starting CLI interface...")
            self.cli_interface.start()
            
        except KeyboardInterrupt:
            print("\n\nShutting down...")
            self.logger.info("CLI application shut down by user")
            
        except Exception as e:
            print(f"\n❌ Unexpected error: {str(e)}")
            self.logger.error(f"Unexpected error in CLI: {str(e)}")
            sys.exit(1)
    
    def get_system_status(self) -> dict:
        """Get current system status for debugging."""
        return {
            'initialized': self.initialized,
            'initialization_errors': self.initialization_errors,
            'llm_connected': self.llm_service.initialize_connection() if self.llm_service else False,
            'template_operations': self.template_registry.get_registry_stats() if self.template_registry else None,
            'excel_loaded': bool(self.excel_service.workbook) if hasattr(self, 'excel_service') else False
        }


def main():
    """Main entry point for the CLI application."""
    # Create and initialize the application
    app = CLIApp()
    
    # Initialize system components
    if not app.initialize():
        print("❌ Failed to initialize system. Check logs for details.")
        sys.exit(1)
    
    # Run the CLI interface
    app.run()


if __name__ == "__main__":
    main()