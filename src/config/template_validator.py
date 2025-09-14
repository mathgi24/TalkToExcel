"""Template configuration validation for Excel-LLM Integration Tool."""

import yaml
from typing import Dict, Any, List, Optional
from pathlib import Path
import importlib.util
import inspect


class TemplateValidator:
    """Validates operation template configurations."""
    
    def __init__(self, template_path: Optional[str] = None):
        """Initialize template validator.
        
        Args:
            template_path: Path to operations template file
        """
        if template_path is None:
            template_path = Path(__file__).parent.parent / "templates" / "operations.yaml"
        
        self.template_path = Path(template_path)
        self._templates: Dict[str, Any] = {}
        self.load_templates()
    
    def load_templates(self) -> None:
        """Load operation templates from YAML file."""
        try:
            if self.template_path.exists():
                with open(self.template_path, 'r', encoding='utf-8') as file:
                    self._templates = yaml.safe_load(file) or {}
            else:
                raise FileNotFoundError(f"Template file not found: {self.template_path}")
        except Exception as e:
            raise RuntimeError(f"Failed to load templates: {e}")
    
    def validate_all_templates(self) -> Dict[str, Any]:
        """Validate all operation templates.
        
        Returns:
            Dict with validation results and any errors found
        """
        errors = []
        warnings = []
        
        if not self._templates:
            errors.append("No templates found in configuration")
            return {'valid': False, 'errors': errors, 'warnings': warnings}
        
        for category_name, category_ops in self._templates.items():
            if not isinstance(category_ops, dict):
                errors.append(f"Category '{category_name}' must be a dictionary")
                continue
            
            for op_name, op_config in category_ops.items():
                op_validation = self._validate_operation(category_name, op_name, op_config)
                errors.extend(op_validation['errors'])
                warnings.extend(op_validation['warnings'])
        
        # Validate for duplicate intent keywords across operations
        keyword_validation = self._validate_intent_keywords()
        warnings.extend(keyword_validation['warnings'])
        
        return {
            'valid': len(errors) == 0,
            'errors': errors,
            'warnings': warnings,
            'total_operations': self._count_total_operations()
        }
    
    def _validate_operation(self, category: str, operation: str, config: Dict[str, Any]) -> Dict[str, List[str]]:
        """Validate a single operation configuration.
        
        Args:
            category: Operation category name
            operation: Operation name
            config: Operation configuration
            
        Returns:
            Dict with validation errors and warnings
        """
        errors = []
        warnings = []
        
        required_fields = ['parameters', 'function', 'safety_level', 'intent_keywords', 'examples', 'description']
        
        # Check required fields
        for field in required_fields:
            if field not in config:
                errors.append(f"Operation '{category}.{operation}' missing required field: {field}")
        
        # Validate parameters
        if 'parameters' in config:
            if not isinstance(config['parameters'], list):
                errors.append(f"Operation '{category}.{operation}' parameters must be a list")
            elif not config['parameters']:
                warnings.append(f"Operation '{category}.{operation}' has no parameters")
        
        # Validate function reference
        if 'function' in config:
            function_validation = self._validate_function_reference(category, operation, config['function'])
            errors.extend(function_validation['errors'])
            warnings.extend(function_validation['warnings'])
        
        # Validate safety level
        if 'safety_level' in config:
            valid_levels = ['safe', 'medium', 'high', 'dangerous']
            if config['safety_level'] not in valid_levels:
                errors.append(f"Operation '{category}.{operation}' has invalid safety_level. Must be one of: {valid_levels}")
        
        # Validate intent keywords
        if 'intent_keywords' in config:
            if not isinstance(config['intent_keywords'], list):
                errors.append(f"Operation '{category}.{operation}' intent_keywords must be a list")
            elif not config['intent_keywords']:
                warnings.append(f"Operation '{category}.{operation}' has no intent keywords")
            else:
                # Check for empty or non-string keywords
                for keyword in config['intent_keywords']:
                    if not isinstance(keyword, str) or not keyword.strip():
                        errors.append(f"Operation '{category}.{operation}' has invalid intent keyword: {keyword}")
        
        # Validate examples
        if 'examples' in config:
            if not isinstance(config['examples'], list):
                errors.append(f"Operation '{category}.{operation}' examples must be a list")
            elif not config['examples']:
                warnings.append(f"Operation '{category}.{operation}' has no examples")
            else:
                # Check for empty or non-string examples
                for example in config['examples']:
                    if not isinstance(example, str) or not example.strip():
                        errors.append(f"Operation '{category}.{operation}' has invalid example: {example}")
        
        # Validate description
        if 'description' in config:
            if not isinstance(config['description'], str) or not config['description'].strip():
                errors.append(f"Operation '{category}.{operation}' description must be a non-empty string")
        
        return {'errors': errors, 'warnings': warnings}
    
    def _validate_function_reference(self, category: str, operation: str, function_ref: str) -> Dict[str, List[str]]:
        """Validate that a function reference points to an existing function.
        
        Args:
            category: Operation category
            operation: Operation name
            function_ref: Function reference string (e.g., 'module.function')
            
        Returns:
            Dict with validation errors and warnings
        """
        errors = []
        warnings = []
        
        if not isinstance(function_ref, str) or '.' not in function_ref:
            errors.append(f"Operation '{category}.{operation}' function reference must be in format 'module.function'")
            return {'errors': errors, 'warnings': warnings}
        
        try:
            module_name, function_name = function_ref.rsplit('.', 1)
            
            # Try to find the module in the src directory
            src_path = Path(__file__).parent.parent
            possible_paths = [
                src_path / f"{module_name.replace('.', '/')}.py",
                src_path / "operations" / f"{module_name}.py",
                src_path / "excel" / f"{module_name}.py",
                src_path / "processing" / f"{module_name}.py"
            ]
            
            module_found = False
            for module_path in possible_paths:
                if module_path.exists():
                    module_found = True
                    # Try to load the module and check if function exists
                    try:
                        spec = importlib.util.spec_from_file_location(module_name, module_path)
                        if spec and spec.loader:
                            module = importlib.util.module_from_spec(spec)
                            spec.loader.exec_module(module)
                            
                            if hasattr(module, function_name):
                                func = getattr(module, function_name)
                                if callable(func):
                                    # Validate function signature if possible
                                    try:
                                        sig = inspect.signature(func)
                                        param_count = len(sig.parameters)
                                        if param_count == 0:
                                            warnings.append(f"Function '{function_ref}' has no parameters")
                                    except Exception:
                                        warnings.append(f"Could not inspect function signature for '{function_ref}'")
                                else:
                                    errors.append(f"'{function_ref}' is not callable")
                            else:
                                errors.append(f"Function '{function_name}' not found in module '{module_name}'")
                    except Exception as e:
                        warnings.append(f"Could not validate function '{function_ref}': {e}")
                    break
            
            if not module_found:
                warnings.append(f"Module '{module_name}' not found for function '{function_ref}'")
                
        except Exception as e:
            errors.append(f"Invalid function reference '{function_ref}': {e}")
        
        return {'errors': errors, 'warnings': warnings}
    
    def _validate_intent_keywords(self) -> Dict[str, List[str]]:
        """Check for duplicate intent keywords across operations.
        
        Returns:
            Dict with validation warnings for duplicate keywords
        """
        warnings = []
        keyword_map = {}
        
        for category_name, category_ops in self._templates.items():
            for op_name, op_config in category_ops.items():
                if 'intent_keywords' in op_config:
                    for keyword in op_config['intent_keywords']:
                        if keyword in keyword_map:
                            warnings.append(
                                f"Intent keyword '{keyword}' is used by both "
                                f"'{keyword_map[keyword]}' and '{category_name}.{op_name}'"
                            )
                        else:
                            keyword_map[keyword] = f"{category_name}.{op_name}"
        
        return {'warnings': warnings}
    
    def _count_total_operations(self) -> int:
        """Count total number of operations across all categories."""
        total = 0
        for category_ops in self._templates.values():
            if isinstance(category_ops, dict):
                total += len(category_ops)
        return total
    
    def get_operation_summary(self) -> Dict[str, Any]:
        """Get a summary of all operations by category.
        
        Returns:
            Dict with operation counts and details by category
        """
        summary = {}
        
        for category_name, category_ops in self._templates.items():
            if isinstance(category_ops, dict):
                operations = []
                for op_name, op_config in category_ops.items():
                    operations.append({
                        'name': op_name,
                        'safety_level': op_config.get('safety_level', 'unknown'),
                        'parameter_count': len(op_config.get('parameters', [])),
                        'keyword_count': len(op_config.get('intent_keywords', [])),
                        'example_count': len(op_config.get('examples', []))
                    })
                
                summary[category_name] = {
                    'count': len(operations),
                    'operations': operations
                }
        
        return summary
    
    def validate_operation_exists(self, category: str, operation: str) -> bool:
        """Check if a specific operation exists in the templates.
        
        Args:
            category: Operation category
            operation: Operation name
            
        Returns:
            True if operation exists
        """
        return (category in self._templates and 
                isinstance(self._templates[category], dict) and
                operation in self._templates[category])
    
    def get_operation_config(self, category: str, operation: str) -> Optional[Dict[str, Any]]:
        """Get configuration for a specific operation.
        
        Args:
            category: Operation category
            operation: Operation name
            
        Returns:
            Operation configuration or None if not found
        """
        if self.validate_operation_exists(category, operation):
            return self._templates[category][operation]
        return None