"""YAML configuration loader for operation definitions."""

import os
import yaml
from typing import Dict, Any, List, Optional
from pathlib import Path
import logging
from datetime import datetime


class TemplateLoader:
    """Loads and manages operation templates from YAML configuration files."""
    
    def __init__(self, templates_dir: Optional[str] = None):
        """Initialize template loader.
        
        Args:
            templates_dir: Directory containing template YAML files. 
                          Defaults to src/templates/
        """
        if templates_dir is None:
            templates_dir = Path(__file__).parent
        
        self.templates_dir = Path(templates_dir)
        self.operations_file = self.templates_dir / "operations.yaml"
        self._operations: Dict[str, Any] = {}
        self._last_modified: Optional[datetime] = None
        self.logger = logging.getLogger(__name__)
        
        # Load operations on initialization
        self.load_operations()
    
    def load_operations(self) -> Dict[str, Any]:
        """Load operation definitions from YAML file.
        
        Returns:
            Dictionary containing all operation definitions
            
        Raises:
            FileNotFoundError: If operations.yaml file is not found
            yaml.YAMLError: If YAML parsing fails
            ValueError: If operation definitions are invalid
        """
        try:
            if not self.operations_file.exists():
                raise FileNotFoundError(f"Operations file not found: {self.operations_file}")
            
            # Check if file has been modified
            current_modified = datetime.fromtimestamp(self.operations_file.stat().st_mtime)
            if self._last_modified and current_modified <= self._last_modified:
                return self._operations
            
            with open(self.operations_file, 'r', encoding='utf-8') as file:
                operations_data = yaml.safe_load(file)
            
            if not operations_data:
                raise ValueError("Operations file is empty or invalid")
            
            # Validate operation structure
            self._validate_operations(operations_data)
            
            self._operations = operations_data
            self._last_modified = current_modified
            
            self.logger.info(f"Loaded {self._count_total_operations()} operations from {len(operations_data)} categories")
            return self._operations
            
        except Exception as e:
            self.logger.error(f"Failed to load operations: {e}")
            raise
    
    def _validate_operations(self, operations_data: Dict[str, Any]) -> None:
        """Validate operation definitions structure.
        
        Args:
            operations_data: Raw operations data from YAML
            
        Raises:
            ValueError: If operation structure is invalid
        """
        required_fields = ['parameters', 'function', 'safety_level', 'intent_keywords', 'examples', 'description']
        valid_safety_levels = ['safe', 'medium', 'high']
        
        for category_name, category_ops in operations_data.items():
            if not isinstance(category_ops, dict):
                raise ValueError(f"Category '{category_name}' must be a dictionary")
            
            for op_name, op_config in category_ops.items():
                if not isinstance(op_config, dict):
                    raise ValueError(f"Operation '{op_name}' in category '{category_name}' must be a dictionary")
                
                # Check required fields
                for field in required_fields:
                    if field not in op_config:
                        raise ValueError(f"Operation '{op_name}' missing required field: {field}")
                
                # Validate safety level
                if op_config['safety_level'] not in valid_safety_levels:
                    raise ValueError(f"Invalid safety level '{op_config['safety_level']}' for operation '{op_name}'")
                
                # Validate parameters is a list
                if not isinstance(op_config['parameters'], list):
                    raise ValueError(f"Parameters for operation '{op_name}' must be a list")
                
                # Validate intent_keywords is a list
                if not isinstance(op_config['intent_keywords'], list):
                    raise ValueError(f"Intent keywords for operation '{op_name}' must be a list")
                
                # Validate examples is a list
                if not isinstance(op_config['examples'], list):
                    raise ValueError(f"Examples for operation '{op_name}' must be a list")
    
    def _count_total_operations(self) -> int:
        """Count total number of operations across all categories."""
        return sum(len(ops) for ops in self._operations.values())
    
    def get_operations(self) -> Dict[str, Any]:
        """Get all operation definitions.
        
        Returns:
            Dictionary containing all operation definitions
        """
        return self._operations.copy()
    
    def get_operation_categories(self) -> List[str]:
        """Get list of operation category names.
        
        Returns:
            List of category names
        """
        return list(self._operations.keys())
    
    def get_operations_by_category(self, category: str) -> Dict[str, Any]:
        """Get operations for a specific category.
        
        Args:
            category: Category name
            
        Returns:
            Dictionary of operations in the category
            
        Raises:
            KeyError: If category does not exist
        """
        if category not in self._operations:
            raise KeyError(f"Category '{category}' not found")
        
        return self._operations[category].copy()
    
    def get_operation(self, category: str, operation: str) -> Dict[str, Any]:
        """Get specific operation definition.
        
        Args:
            category: Category name
            operation: Operation name
            
        Returns:
            Operation definition dictionary
            
        Raises:
            KeyError: If category or operation does not exist
        """
        if category not in self._operations:
            raise KeyError(f"Category '{category}' not found")
        
        if operation not in self._operations[category]:
            raise KeyError(f"Operation '{operation}' not found in category '{category}'")
        
        return self._operations[category][operation].copy()
    
    def get_all_intent_keywords(self) -> List[str]:
        """Get all intent keywords from all operations.
        
        Returns:
            Flattened list of all intent keywords
        """
        keywords = []
        for category_ops in self._operations.values():
            for op_config in category_ops.values():
                keywords.extend(op_config.get('intent_keywords', []))
        return list(set(keywords))  # Remove duplicates
    
    def get_operations_by_safety_level(self, safety_level: str) -> Dict[str, Dict[str, Any]]:
        """Get operations filtered by safety level.
        
        Args:
            safety_level: Safety level to filter by ('safe', 'medium', 'high')
            
        Returns:
            Dictionary of operations grouped by category
        """
        filtered_ops = {}
        for category_name, category_ops in self._operations.items():
            category_filtered = {}
            for op_name, op_config in category_ops.items():
                if op_config.get('safety_level') == safety_level:
                    category_filtered[op_name] = op_config
            
            if category_filtered:
                filtered_ops[category_name] = category_filtered
        
        return filtered_ops
    
    def search_operations_by_keyword(self, keyword: str) -> Dict[str, Dict[str, Any]]:
        """Search operations by intent keyword.
        
        Args:
            keyword: Keyword to search for
            
        Returns:
            Dictionary of matching operations grouped by category
        """
        matching_ops = {}
        keyword_lower = keyword.lower()
        
        for category_name, category_ops in self._operations.items():
            category_matches = {}
            for op_name, op_config in category_ops.items():
                intent_keywords = [kw.lower() for kw in op_config.get('intent_keywords', [])]
                if keyword_lower in intent_keywords:
                    category_matches[op_name] = op_config
            
            if category_matches:
                matching_ops[category_name] = category_matches
        
        return matching_ops
    
    def reload_operations(self) -> Dict[str, Any]:
        """Force reload operations from file.
        
        Returns:
            Updated operations dictionary
        """
        self._last_modified = None  # Force reload
        return self.load_operations()
    
    def is_file_modified(self) -> bool:
        """Check if operations file has been modified since last load.
        
        Returns:
            True if file has been modified
        """
        if not self.operations_file.exists():
            return False
        
        current_modified = datetime.fromtimestamp(self.operations_file.stat().st_mtime)
        return self._last_modified is None or current_modified > self._last_modified