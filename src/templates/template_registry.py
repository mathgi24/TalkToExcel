"""Template registry that maps operation names to functions."""

import importlib
import inspect
import sys
import os
from typing import Dict, Any, Callable, Optional, List
from pathlib import Path
import logging

from .template_loader import TemplateLoader


class TemplateRegistry:
    """Registry that maps operation names to their corresponding functions."""
    
    def __init__(self, template_loader: Optional[TemplateLoader] = None):
        """Initialize template registry.
        
        Args:
            template_loader: TemplateLoader instance. If None, creates a new one.
        """
        self.template_loader = template_loader or TemplateLoader()
        self._function_registry: Dict[str, Callable] = {}
        self._operation_metadata: Dict[str, Dict[str, Any]] = {}
        self.logger = logging.getLogger(__name__)
        
        # Ensure src directory is in Python path for imports
        self._ensure_src_in_path()
        
        # Initialize registry
        self._build_registry()
    
    def _ensure_src_in_path(self) -> None:
        """Ensure src directory is in Python path for imports."""
        # Get the src directory path
        current_file = Path(__file__)
        src_dir = current_file.parent.parent  # Go up from templates to src
        src_path = str(src_dir)
        
        # Add to Python path if not already there
        if src_path not in sys.path:
            sys.path.insert(0, src_path)
    
    def _build_registry(self) -> None:
        """Build the function registry from operation definitions."""
        operations = self.template_loader.get_operations()
        
        for category_name, category_ops in operations.items():
            for op_name, op_config in category_ops.items():
                function_path = op_config.get('function')
                if function_path:
                    registry_key = f"{category_name}.{op_name}"
                    
                    try:
                        # Load and register the function
                        function = self._load_function(function_path)
                        self._function_registry[registry_key] = function
                        
                        # Store metadata
                        self._operation_metadata[registry_key] = {
                            'category': category_name,
                            'operation': op_name,
                            'config': op_config,
                            'function_path': function_path
                        }
                        
                    except Exception as e:
                        self.logger.warning(f"Failed to load function for {registry_key}: {e}")
                        # Create a placeholder function for missing implementations
                        self._function_registry[registry_key] = self._create_placeholder_function(
                            category_name, op_name, function_path
                        )
                        self._operation_metadata[registry_key] = {
                            'category': category_name,
                            'operation': op_name,
                            'config': op_config,
                            'function_path': function_path,
                            'placeholder': True
                        }
        
        self.logger.info(f"Registry built with {len(self._function_registry)} operations")
    
    def _load_function(self, function_path: str) -> Callable:
        """Load a function from its module path.
        
        Args:
            function_path: Function path in format 'module.function_name'
            
        Returns:
            The loaded function
            
        Raises:
            ImportError: If module cannot be imported
            AttributeError: If function is not found in module
        """
        if '.' not in function_path:
            raise ValueError(f"Invalid function path format: {function_path}")
        
        module_path, function_name = function_path.rsplit('.', 1)
        
        try:
            # Try to import from operations package with src prefix
            full_module_path = f"src.operations.{module_path}"
            module = importlib.import_module(full_module_path)
        except ImportError:
            try:
                # Fallback: try without src prefix
                full_module_path = f"operations.{module_path}"
                module = importlib.import_module(full_module_path)
            except ImportError:
                # Fallback to direct import
                try:
                    module = importlib.import_module(module_path)
                except ImportError as e:
                    raise ImportError(f"Cannot import module '{module_path}': {e}")
        
        if not hasattr(module, function_name):
            raise AttributeError(f"Function '{function_name}' not found in module '{module_path}'")
        
        function = getattr(module, function_name)
        
        if not callable(function):
            raise TypeError(f"'{function_path}' is not callable")
        
        return function
    
    def _create_placeholder_function(self, category: str, operation: str, function_path: str) -> Callable:
        """Create a placeholder function for missing implementations.
        
        Args:
            category: Operation category
            operation: Operation name
            function_path: Expected function path
            
        Returns:
            Placeholder function that raises NotImplementedError
        """
        def placeholder(*args, **kwargs):
            raise NotImplementedError(
                f"Operation '{category}.{operation}' is not implemented. "
                f"Expected function at: {function_path}"
            )
        
        placeholder.__name__ = f"{category}_{operation}_placeholder"
        placeholder.__doc__ = f"Placeholder for {category}.{operation} operation"
        
        return placeholder
    
    def get_function(self, category: str, operation: str) -> Optional[Callable]:
        """Get function for a specific operation.
        
        Args:
            category: Operation category
            operation: Operation name
            
        Returns:
            Function callable or None if not found
        """
        registry_key = f"{category}.{operation}"
        return self._function_registry.get(registry_key)
    
    def get_function_by_key(self, registry_key: str) -> Optional[Callable]:
        """Get function by registry key.
        
        Args:
            registry_key: Registry key in format 'category.operation'
            
        Returns:
            Function callable or None if not found
        """
        return self._function_registry.get(registry_key)
    
    def execute_operation(self, category: str, operation: str, *args, **kwargs) -> Any:
        """Execute an operation with given parameters.
        
        Args:
            category: Operation category
            operation: Operation name
            *args: Positional arguments for the operation
            **kwargs: Keyword arguments for the operation
            
        Returns:
            Operation result
            
        Raises:
            KeyError: If operation is not found
            Exception: Any exception raised by the operation function
        """
        function = self.get_function(category, operation)
        if function is None:
            raise KeyError(f"Operation '{category}.{operation}' not found in registry")
        
        try:
            return function(*args, **kwargs)
        except Exception as e:
            self.logger.error(f"Error executing {category}.{operation}: {e}")
            raise
    
    def get_operation_metadata(self, category: str, operation: str) -> Optional[Dict[str, Any]]:
        """Get metadata for a specific operation.
        
        Args:
            category: Operation category
            operation: Operation name
            
        Returns:
            Operation metadata dictionary or None if not found
        """
        registry_key = f"{category}.{operation}"
        return self._operation_metadata.get(registry_key)
    
    def get_all_operations(self) -> List[str]:
        """Get list of all registered operation keys.
        
        Returns:
            List of operation keys in format 'category.operation'
        """
        return list(self._function_registry.keys())
    
    def get_operations_by_category(self, category: str) -> List[str]:
        """Get operations for a specific category.
        
        Args:
            category: Category name
            
        Returns:
            List of operation names in the category
        """
        operations = []
        for key in self._function_registry.keys():
            if key.startswith(f"{category}."):
                operations.append(key.split('.', 1)[1])
        return operations
    
    def is_operation_available(self, category: str, operation: str) -> bool:
        """Check if an operation is available and implemented.
        
        Args:
            category: Operation category
            operation: Operation name
            
        Returns:
            True if operation is available and not a placeholder
        """
        registry_key = f"{category}.{operation}"
        metadata = self._operation_metadata.get(registry_key)
        
        if metadata is None:
            return False
        
        return not metadata.get('placeholder', False)
    
    def get_function_signature(self, category: str, operation: str) -> Optional[inspect.Signature]:
        """Get function signature for an operation.
        
        Args:
            category: Operation category
            operation: Operation name
            
        Returns:
            Function signature or None if not found
        """
        function = self.get_function(category, operation)
        if function is None:
            return None
        
        try:
            return inspect.signature(function)
        except (ValueError, TypeError):
            return None
    
    def validate_operation_parameters(self, category: str, operation: str, 
                                    args: tuple, kwargs: dict) -> bool:
        """Validate parameters for an operation.
        
        Args:
            category: Operation category
            operation: Operation name
            args: Positional arguments
            kwargs: Keyword arguments
            
        Returns:
            True if parameters are valid
        """
        signature = self.get_function_signature(category, operation)
        if signature is None:
            return False
        
        try:
            signature.bind(*args, **kwargs)
            return True
        except TypeError:
            return False
    
    def reload_registry(self) -> None:
        """Reload the registry from updated operation definitions."""
        self.template_loader.reload_operations()
        self._function_registry.clear()
        self._operation_metadata.clear()
        self._build_registry()
    
    def get_all_operations_with_examples(self) -> Dict[str, Dict[str, Any]]:
        """Get all operations with their examples and metadata for LLM service.
        
        Returns:
            Dictionary with operation categories and their configurations
        """
        operations_data = self.template_loader.get_operations()
        
        # Transform the data for LLM service consumption
        transformed_operations = {}
        
        for category_name, category_ops in operations_data.items():
            transformed_operations[category_name] = {}
            
            for op_name, op_config in category_ops.items():
                transformed_operations[category_name][op_name] = {
                    'parameters': op_config.get('parameters', []),
                    'safety_level': op_config.get('safety_level', 'safe'),
                    'intent_keywords': op_config.get('intent_keywords', []),
                    'examples': op_config.get('examples', []),
                    'description': op_config.get('description', '')
                }
        
        return transformed_operations
    
        self.logger.info("Registry reloaded successfully")
    
    def get_registry_stats(self) -> Dict[str, Any]:
        """Get statistics about the registry.
        
        Returns:
            Dictionary with registry statistics
        """
        total_operations = len(self._function_registry)
        placeholder_count = sum(
            1 for metadata in self._operation_metadata.values() 
            if metadata.get('placeholder', False)
        )
        implemented_count = total_operations - placeholder_count
        
        categories = set()
        for key in self._function_registry.keys():
            categories.add(key.split('.', 1)[0])
        
        return {
            'total_operations': total_operations,
            'implemented_operations': implemented_count,
            'placeholder_operations': placeholder_count,
            'categories': len(categories),
            'category_names': sorted(list(categories))
        }
    
    def is_loaded(self) -> bool:
        """Check if the template registry is loaded with operations.
        
        Returns:
            True if registry has operations loaded, False otherwise
        """
        return len(self._function_registry) > 0
    
    def cleanup(self):
        """Clean up resources used by the registry."""
        self._function_registry.clear()
        self._operation_metadata.clear()
        if hasattr(self.template_loader, 'cleanup'):
            self.template_loader.cleanup()