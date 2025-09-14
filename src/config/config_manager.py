"""Configuration management for Excel-LLM Integration Tool."""

import os
import yaml
import shutil
from typing import Dict, Any, Optional, List
from pathlib import Path
import requests
from urllib.parse import urlparse


class ConfigManager:
    """Manages application configuration from YAML files."""
    
    def __init__(self, config_path: Optional[str] = None):
        """Initialize configuration manager.
        
        Args:
            config_path: Path to configuration file. Defaults to config/config.yaml
        """
        if config_path is None:
            config_path = Path(__file__).parent.parent.parent / "config" / "config.yaml"
        
        self.config_path = Path(config_path)
        self._config: Dict[str, Any] = {}
        self.load_config()
    
    def load_config(self) -> None:
        """Load configuration from YAML file."""
        try:
            if self.config_path.exists():
                with open(self.config_path, 'r', encoding='utf-8') as file:
                    self._config = yaml.safe_load(file) or {}
            else:
                raise FileNotFoundError(f"Configuration file not found: {self.config_path}")
        except Exception as e:
            raise RuntimeError(f"Failed to load configuration: {e}")
    
    def get(self, key: str, default: Any = None) -> Any:
        """Get configuration value by key.
        
        Args:
            key: Configuration key (supports dot notation, e.g., 'ollama.endpoint')
            default: Default value if key not found
            
        Returns:
            Configuration value or default
        """
        keys = key.split('.')
        value = self._config
        
        for k in keys:
            if isinstance(value, dict) and k in value:
                value = value[k]
            else:
                return default
        
        return value
    
    def get_ollama_config(self) -> Dict[str, Any]:
        """Get Ollama configuration."""
        return self.get('ollama', {})
    
    def get_backup_config(self) -> Dict[str, Any]:
        """Get backup configuration."""
        return self.get('backup', {})
    
    def get_safety_config(self) -> Dict[str, Any]:
        """Get safety configuration."""
        return self.get('safety', {})
    
    def get_excel_config(self) -> Dict[str, Any]:
        """Get Excel configuration."""
        return self.get('excel', {})
    
    def get_logging_config(self) -> Dict[str, Any]:
        """Get logging configuration."""
        return self.get('logging', {})
    
    def validate_config(self) -> bool:
        """Validate configuration completeness and correctness.
        
        Returns:
            True if configuration is valid
            
        Raises:
            ValueError: If configuration is invalid
        """
        validation_results = self.validate_all_sections()
        
        if not validation_results['valid']:
            raise ValueError(f"Configuration validation failed: {validation_results['errors']}")
        
        return True
    
    def validate_all_sections(self) -> Dict[str, Any]:
        """Comprehensive validation of all configuration sections.
        
        Returns:
            Dict with validation results and any errors found
        """
        errors = []
        warnings = []
        
        # Validate required keys
        required_keys = [
            'ollama.endpoint',
            'ollama.model',
            'backup.directory',
            'safety.max_rows_per_operation'
        ]
        
        for key in required_keys:
            if self.get(key) is None:
                errors.append(f"Required configuration key missing: {key}")
        
        # Validate Ollama configuration
        ollama_validation = self._validate_ollama_config()
        errors.extend(ollama_validation['errors'])
        warnings.extend(ollama_validation['warnings'])
        
        # Validate backup configuration
        backup_validation = self._validate_backup_config()
        errors.extend(backup_validation['errors'])
        warnings.extend(backup_validation['warnings'])
        
        # Validate safety configuration
        safety_validation = self._validate_safety_config()
        errors.extend(safety_validation['errors'])
        warnings.extend(safety_validation['warnings'])
        
        # Validate Excel configuration
        excel_validation = self._validate_excel_config()
        errors.extend(excel_validation['errors'])
        warnings.extend(excel_validation['warnings'])
        
        # Validate logging configuration
        logging_validation = self._validate_logging_config()
        errors.extend(logging_validation['errors'])
        warnings.extend(logging_validation['warnings'])
        
        return {
            'valid': len(errors) == 0,
            'errors': errors,
            'warnings': warnings
        }
    
    def _validate_ollama_config(self) -> Dict[str, List[str]]:
        """Validate Ollama configuration section."""
        errors = []
        warnings = []
        
        endpoint = self.get('ollama.endpoint')
        if endpoint:
            # Validate endpoint format
            if not endpoint.startswith(('http://', 'https://')):
                errors.append(f"Invalid Ollama endpoint format: {endpoint}")
            else:
                # Validate endpoint accessibility
                try:
                    parsed = urlparse(endpoint)
                    if not parsed.hostname:
                        errors.append(f"Invalid Ollama endpoint hostname: {endpoint}")
                except Exception as e:
                    errors.append(f"Error parsing Ollama endpoint: {e}")
        
        # Validate model name
        model = self.get('ollama.model')
        if model and not isinstance(model, str):
            errors.append("Ollama model must be a string")
        
        # Validate temperature
        temperature = self.get('ollama.temperature')
        if temperature is not None:
            if not isinstance(temperature, (int, float)) or not 0 <= temperature <= 2:
                errors.append("Ollama temperature must be a number between 0 and 2")
        
        # Validate max_tokens
        max_tokens = self.get('ollama.max_tokens')
        if max_tokens is not None:
            if not isinstance(max_tokens, int) or max_tokens <= 0:
                errors.append("Ollama max_tokens must be a positive integer")
        
        # Validate timeout
        timeout = self.get('ollama.timeout')
        if timeout is not None:
            if not isinstance(timeout, (int, float)) or timeout <= 0:
                errors.append("Ollama timeout must be a positive number")
        
        # Validate retry settings
        retry_attempts = self.get('ollama.retry_attempts')
        if retry_attempts is not None:
            if not isinstance(retry_attempts, int) or retry_attempts < 0:
                errors.append("Ollama retry_attempts must be a non-negative integer")
        
        retry_delay = self.get('ollama.retry_delay')
        if retry_delay is not None:
            if not isinstance(retry_delay, (int, float)) or retry_delay < 0:
                errors.append("Ollama retry_delay must be a non-negative number")
        
        return {'errors': errors, 'warnings': warnings}
    
    def _validate_backup_config(self) -> Dict[str, List[str]]:
        """Validate backup configuration section."""
        errors = []
        warnings = []
        
        # Validate backup directory
        backup_dir = self.get('backup.directory')
        if backup_dir:
            try:
                backup_path = Path(backup_dir)
                backup_path.mkdir(parents=True, exist_ok=True)
                
                # Check if directory is writable
                test_file = backup_path / '.test_write'
                try:
                    test_file.touch()
                    test_file.unlink()
                except Exception:
                    errors.append(f"Backup directory is not writable: {backup_dir}")
            except Exception as e:
                errors.append(f"Cannot create backup directory: {e}")
        
        # Validate retention count
        retention_count = self.get('backup.retention_count')
        if retention_count is not None:
            if not isinstance(retention_count, int) or retention_count < 1:
                errors.append("Backup retention_count must be a positive integer")
        
        # Validate timestamp format
        timestamp_format = self.get('backup.timestamp_format')
        if timestamp_format:
            try:
                from datetime import datetime
                datetime.now().strftime(timestamp_format)
            except Exception as e:
                errors.append(f"Invalid backup timestamp format: {e}")
        
        return {'errors': errors, 'warnings': warnings}
    
    def _validate_safety_config(self) -> Dict[str, List[str]]:
        """Validate safety configuration section."""
        errors = []
        warnings = []
        
        # Validate max rows per operation
        max_rows = self.get('safety.max_rows_per_operation')
        if max_rows is not None:
            if not isinstance(max_rows, int) or max_rows <= 0:
                errors.append("Safety max_rows_per_operation must be a positive integer")
            elif max_rows > 1000:
                warnings.append(f"Large max_rows_per_operation ({max_rows}) may impact performance")
        
        # Validate max columns per operation
        max_columns = self.get('safety.max_columns_per_operation')
        if max_columns is not None:
            if not isinstance(max_columns, int) or max_columns <= 0:
                errors.append("Safety max_columns_per_operation must be a positive integer")
        
        return {'errors': errors, 'warnings': warnings}
    
    def _validate_excel_config(self) -> Dict[str, List[str]]:
        """Validate Excel configuration section."""
        errors = []
        warnings = []
        
        # Validate supported formats
        supported_formats = self.get('excel.supported_formats')
        if supported_formats:
            if not isinstance(supported_formats, list):
                errors.append("Excel supported_formats must be a list")
            else:
                valid_formats = ['.xlsx', '.xls', '.csv', '.xlsm', '.xlsb']
                for fmt in supported_formats:
                    if fmt not in valid_formats:
                        warnings.append(f"Unsupported Excel format: {fmt}")
        
        return {'errors': errors, 'warnings': warnings}
    
    def _validate_logging_config(self) -> Dict[str, List[str]]:
        """Validate logging configuration section."""
        errors = []
        warnings = []
        
        # Validate log level
        log_level = self.get('logging.level')
        if log_level:
            valid_levels = ['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL']
            if log_level not in valid_levels:
                errors.append(f"Invalid logging level: {log_level}. Must be one of {valid_levels}")
        
        # Validate log file path
        log_file = self.get('logging.file')
        if log_file:
            try:
                log_path = Path(log_file)
                log_path.parent.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                errors.append(f"Cannot create log directory: {e}")
        
        # Validate max file size
        max_size = self.get('logging.max_file_size')
        if max_size and isinstance(max_size, str):
            if not max_size.endswith(('KB', 'MB', 'GB')):
                errors.append("Logging max_file_size must end with KB, MB, or GB")
        
        # Validate backup count
        backup_count = self.get('logging.backup_count')
        if backup_count is not None:
            if not isinstance(backup_count, int) or backup_count < 0:
                errors.append("Logging backup_count must be a non-negative integer")
        
        return {'errors': errors, 'warnings': warnings}
    
    def test_ollama_connection(self) -> Dict[str, Any]:
        """Test connection to Ollama service.
        
        Returns:
            Dict with connection test results
        """
        endpoint = self.get('ollama.endpoint')
        if not endpoint:
            return {'success': False, 'error': 'No Ollama endpoint configured'}
        
        try:
            # Test basic connectivity
            response = requests.get(f"{endpoint}/api/tags", timeout=5)
            if response.status_code == 200:
                models = response.json().get('models', [])
                configured_model = self.get('ollama.model')
                
                model_available = any(
                    model.get('name') == configured_model 
                    for model in models
                )
                
                return {
                    'success': True,
                    'endpoint_accessible': True,
                    'models_available': [model.get('name') for model in models],
                    'configured_model_available': model_available,
                    'configured_model': configured_model
                }
            else:
                return {
                    'success': False,
                    'error': f'Ollama service returned status {response.status_code}'
                }
        except requests.exceptions.ConnectionError:
            return {
                'success': False,
                'error': 'Cannot connect to Ollama service. Is it running?'
            }
        except requests.exceptions.Timeout:
            return {
                'success': False,
                'error': 'Connection to Ollama service timed out'
            }
        except Exception as e:
            return {
                'success': False,
                'error': f'Unexpected error testing Ollama connection: {e}'
            }
    
    def create_default_config(self, config_path: Optional[Path] = None) -> None:
        """Create a default configuration file.
        
        Args:
            config_path: Path where to create the config file
        """
        if config_path is None:
            config_path = self.config_path
        
        default_config = {
            'ollama': {
                'endpoint': 'http://localhost:11434',
                'model': 'mistral:7b-instruct',
                'temperature': 0.1,
                'max_tokens': 1000,
                'timeout': 30,
                'retry_attempts': 3,
                'retry_delay': 2
            },
            'backup': {
                'enabled': True,
                'directory': './backups',
                'retention_count': 10,
                'timestamp_format': '%Y%m%d_%H%M%S',
                'auto_backup': True
            },
            'safety': {
                'max_rows_per_operation': 50,
                'max_columns_per_operation': 20,
                'dangerous_commands_blocked': True,
                'confirmation_required_for_deletes': True
            },
            'excel': {
                'supported_formats': ['.xlsx', '.xls', '.csv'],
                'default_sheet_name': 'Sheet1',
                'auto_detect_headers': True,
                'auto_detect_data_types': True
            },
            'logging': {
                'level': 'INFO',
                'file': './logs/excel_llm.log',
                'max_file_size': '10MB',
                'backup_count': 5
            }
        }
        
        config_path.parent.mkdir(parents=True, exist_ok=True)
        with open(config_path, 'w', encoding='utf-8') as file:
            yaml.dump(default_config, file, default_flow_style=False, indent=2)
    
    def update_config(self, key: str, value: Any) -> None:
        """Update a configuration value and save to file.
        
        Args:
            key: Configuration key (supports dot notation)
            value: New value to set
        """
        keys = key.split('.')
        config = self._config
        
        # Navigate to the parent of the target key
        for k in keys[:-1]:
            if k not in config:
                config[k] = {}
            config = config[k]
        
        # Set the value
        config[keys[-1]] = value
        
        # Save to file
        with open(self.config_path, 'w', encoding='utf-8') as file:
            yaml.dump(self._config, file, default_flow_style=False, indent=2)


# Global configuration instance
config = ConfigManager()