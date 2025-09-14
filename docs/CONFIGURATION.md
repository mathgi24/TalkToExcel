# Configuration Guide

This document provides comprehensive information about configuring the Excel-LLM Integration Tool.

## Table of Contents

- [Overview](#overview)
- [Configuration File Structure](#configuration-file-structure)
- [Configuration Sections](#configuration-sections)
- [Setup and Validation](#setup-and-validation)
- [Operation Templates](#operation-templates)
- [Troubleshooting](#troubleshooting)
- [Examples](#examples)

## Overview

The Excel-LLM Integration Tool uses YAML configuration files to manage system settings. The main configuration file is located at `config/config.yaml` and contains settings for:

- Ollama LLM service connection
- Backup and file management
- Safety and security settings
- Excel file handling
- Logging configuration

## Configuration File Structure

The configuration file is organized into the following main sections:

```yaml
ollama:          # LLM service configuration
backup:          # Backup and file management
safety:          # Safety and security settings
excel:           # Excel file handling
logging:         # Logging configuration
```

## Configuration Sections

### Ollama Configuration

Controls connection to the local Ollama LLM service.

```yaml
ollama:
  endpoint: "http://localhost:11434"    # Ollama service endpoint
  model: "mistral:7b-instruct"          # Model to use for processing
  temperature: 0.1                      # Response randomness (0.0-2.0)
  max_tokens: 1000                      # Maximum response length
  timeout: 30                           # Connection timeout in seconds
  retry_attempts: 3                     # Number of retry attempts
  retry_delay: 2                        # Delay between retries in seconds
```

**Configuration Details:**

- **endpoint**: URL where Ollama service is running
  - Default: `http://localhost:11434`
  - Must start with `http://` or `https://`
  - Validate accessibility during setup

- **model**: Ollama model name to use
  - Default: `mistral:7b-instruct`
  - Must be pulled/available in Ollama
  - Check with: `ollama list`

- **temperature**: Controls response randomness
  - Range: 0.0 (deterministic) to 2.0 (very random)
  - Recommended: 0.1 for consistent command parsing

- **max_tokens**: Maximum response length
  - Default: 1000
  - Adjust based on complexity of operations

- **timeout**: Connection timeout in seconds
  - Default: 30
  - Increase for slower systems

- **retry_attempts**: Number of retry attempts on failure
  - Default: 3
  - Set to 0 to disable retries

- **retry_delay**: Delay between retry attempts
  - Default: 2 seconds
  - Exponential backoff is applied

### Backup Configuration

Controls automatic backup creation and management.

```yaml
backup:
  enabled: true                         # Enable/disable backups
  directory: "./backups"                # Backup storage directory
  retention_count: 10                   # Number of backups to keep
  timestamp_format: "%Y%m%d_%H%M%S"     # Timestamp format for backup files
  auto_backup: true                     # Automatic backup before operations
```

**Configuration Details:**

- **enabled**: Master switch for backup functionality
  - Default: `true`
  - Strongly recommended to keep enabled

- **directory**: Where backup files are stored
  - Default: `./backups`
  - Directory is created automatically
  - Must be writable

- **retention_count**: Number of backup files to keep
  - Default: 10
  - Older backups are automatically deleted
  - Set to 0 for unlimited retention

- **timestamp_format**: Python strftime format for backup filenames
  - Default: `%Y%m%d_%H%M%S` (e.g., 20240315_143022)
  - Must be valid strftime format

- **auto_backup**: Create backup before every operation
  - Default: `true`
  - Applies to all operations, not just destructive ones

### Safety Configuration

Controls safety mechanisms and operation limits.

```yaml
safety:
  max_rows_per_operation: 50            # Maximum rows affected per operation
  max_columns_per_operation: 20         # Maximum columns affected per operation
  dangerous_commands_blocked: true      # Block dangerous operations
  confirmation_required_for_deletes: true  # Require confirmation for deletions
```

**Configuration Details:**

- **max_rows_per_operation**: Hard limit on rows affected
  - Default: 50
  - Operations exceeding this limit are blocked
  - Prevents accidental mass operations

- **max_columns_per_operation**: Hard limit on columns affected
  - Default: 20
  - Similar protection for column operations

- **dangerous_commands_blocked**: Enable dangerous command blocking
  - Default: `true`
  - Blocks operations like "format all", "delete everything"
  - Cannot be overridden by user

- **confirmation_required_for_deletes**: Require user confirmation
  - Default: `true`
  - Applies to all delete operations
  - Provides additional safety layer

### Excel Configuration

Controls Excel file handling and format support.

```yaml
excel:
  supported_formats: [".xlsx", ".xls", ".csv"]  # Supported file formats
  default_sheet_name: "Sheet1"                  # Default sheet name
  auto_detect_headers: true                     # Automatically detect headers
  auto_detect_data_types: true                  # Automatically detect data types
```

**Configuration Details:**

- **supported_formats**: List of supported file extensions
  - Default: `[".xlsx", ".xls", ".csv"]`
  - Additional formats: `.xlsm`, `.xlsb`
  - Files with unsupported formats are rejected

- **default_sheet_name**: Default sheet to use when not specified
  - Default: `"Sheet1"`
  - Used when user doesn't specify sheet name

- **auto_detect_headers**: Automatically detect header rows
  - Default: `true`
  - Improves data interpretation accuracy

- **auto_detect_data_types**: Automatically detect column data types
  - Default: `true`
  - Enables better data validation

### Logging Configuration

Controls application logging behavior.

```yaml
logging:
  level: "INFO"                         # Log level
  file: "./logs/excel_llm.log"          # Log file path
  max_file_size: "10MB"                 # Maximum log file size
  backup_count: 5                       # Number of log backup files
```

**Configuration Details:**

- **level**: Logging level
  - Options: `DEBUG`, `INFO`, `WARNING`, `ERROR`, `CRITICAL`
  - Default: `INFO`
  - `DEBUG` provides detailed operation information

- **file**: Path to log file
  - Default: `./logs/excel_llm.log`
  - Directory is created automatically

- **max_file_size**: Maximum size before log rotation
  - Default: `10MB`
  - Supports: `KB`, `MB`, `GB` suffixes

- **backup_count**: Number of rotated log files to keep
  - Default: 5
  - Older log files are automatically deleted

## Setup and Validation

### Initial Setup

Run the setup script to create and validate configuration:

```bash
# Interactive setup (recommended for first time)
python setup_system.py

# Non-interactive setup
python setup_system.py --non-interactive

# Configuration only
python setup_system.py --config-only

# Validate existing configuration
python setup_system.py --validate-only
```

### Manual Configuration

1. **Create configuration file**:
   ```bash
   mkdir -p config
   cp config/config.yaml.example config/config.yaml
   ```

2. **Edit configuration**:
   ```bash
   # Edit with your preferred editor
   nano config/config.yaml
   ```

3. **Validate configuration**:
   ```python
   from src.config.config_manager import ConfigManager
   
   config = ConfigManager()
   validation_result = config.validate_all_sections()
   
   if validation_result['valid']:
       print("Configuration is valid")
   else:
       print("Errors:", validation_result['errors'])
   ```

### Configuration Validation

The system performs comprehensive validation including:

- **Required fields**: Ensures all mandatory settings are present
- **Data types**: Validates that values are correct types
- **Value ranges**: Checks that numeric values are within valid ranges
- **File paths**: Verifies that directories exist and are writable
- **Network endpoints**: Tests connectivity to Ollama service
- **Format validation**: Ensures formats like timestamps are valid

## Operation Templates

Operation templates are configured in `src/templates/operations.yaml` and define available natural language operations.

### Template Structure

```yaml
category_name:
  operation_name:
    parameters: ["param1", "param2"]      # Function parameters
    function: "module.function_name"      # Implementation function
    safety_level: "safe"                  # Safety classification
    intent_keywords: ["keyword1", "keyword2"]  # Recognition keywords
    examples:                             # Example phrases
      - "example command 1"
      - "example command 2"
    description: "Operation description"  # Human-readable description
```

### Template Validation

Templates are automatically validated for:

- **Required fields**: All mandatory fields present
- **Function references**: Implementation functions exist
- **Safety levels**: Valid safety classifications
- **Intent keywords**: Non-empty keyword lists
- **Examples**: Valid example phrases
- **Duplicate detection**: No conflicting keywords

### Adding New Operations

1. **Define operation in templates**:
   ```yaml
   data_operations:
     my_new_operation:
       parameters: ["sheet_name", "custom_param"]
       function: "data_operations.my_new_function"
       safety_level: "medium"
       intent_keywords: ["custom", "special", "my operation"]
       examples:
         - "perform my custom operation"
         - "apply special function to data"
       description: "My custom data operation"
   ```

2. **Implement function**:
   ```python
   # In src/operations/data_operations.py
   def my_new_function(sheet_name: str, custom_param: str):
       # Implementation here
       pass
   ```

3. **Validate templates**:
   ```bash
   python setup_system.py --validate-only
   ```

## Troubleshooting

### Common Configuration Issues

1. **Ollama Connection Failed**
   ```
   Error: Cannot connect to Ollama service
   ```
   **Solutions**:
   - Start Ollama service: `ollama serve`
   - Check endpoint URL in configuration
   - Verify firewall settings
   - Test manually: `curl http://localhost:11434/api/tags`

2. **Model Not Available**
   ```
   Error: Configured model 'mistral:7b-instruct' not found
   ```
   **Solutions**:
   - Pull model: `ollama pull mistral:7b-instruct`
   - Check available models: `ollama list`
   - Update configuration with available model

3. **Backup Directory Issues**
   ```
   Error: Backup directory is not writable
   ```
   **Solutions**:
   - Check directory permissions
   - Create directory manually: `mkdir -p backups`
   - Change backup directory in configuration

4. **Template Validation Errors**
   ```
   Error: Function 'module.function' not found
   ```
   **Solutions**:
   - Check function implementation exists
   - Verify module path is correct
   - Ensure function is properly imported

### Configuration Reset

To reset configuration to defaults:

```bash
# Backup current configuration
cp config/config.yaml config/config.yaml.backup

# Create new default configuration
python setup_system.py --config-only --non-interactive
```

### Debug Mode

Enable debug logging for troubleshooting:

```yaml
logging:
  level: "DEBUG"
```

This provides detailed information about:
- Configuration loading
- Template validation
- Ollama communication
- Operation execution
- Error details

## Examples

### Development Configuration

For development with detailed logging:

```yaml
ollama:
  endpoint: "http://localhost:11434"
  model: "mistral:7b-instruct"
  temperature: 0.0  # Deterministic responses
  timeout: 60       # Longer timeout for debugging

backup:
  enabled: true
  retention_count: 20  # Keep more backups

safety:
  max_rows_per_operation: 10  # Stricter limits for testing

logging:
  level: "DEBUG"    # Detailed logging
  max_file_size: "50MB"  # Larger log files
```

### Production Configuration

For production use with enhanced safety:

```yaml
ollama:
  endpoint: "http://localhost:11434"
  model: "mistral:7b-instruct"
  temperature: 0.1
  retry_attempts: 5  # More retries

backup:
  enabled: true
  retention_count: 50  # Keep more backups
  auto_backup: true

safety:
  max_rows_per_operation: 25  # Conservative limits
  dangerous_commands_blocked: true
  confirmation_required_for_deletes: true

logging:
  level: "INFO"
  backup_count: 10  # Keep more log history
```

### Minimal Configuration

For basic usage with minimal features:

```yaml
ollama:
  endpoint: "http://localhost:11434"
  model: "mistral:7b-instruct"

backup:
  enabled: true
  directory: "./backups"

safety:
  max_rows_per_operation: 50

excel:
  supported_formats: [".xlsx"]

logging:
  level: "WARNING"  # Minimal logging
```

## Configuration API

### Programmatic Access

```python
from src.config.config_manager import ConfigManager

# Load configuration
config = ConfigManager()

# Get specific values
endpoint = config.get('ollama.endpoint')
backup_dir = config.get('backup.directory', './default_backups')

# Get section configurations
ollama_config = config.get_ollama_config()
safety_config = config.get_safety_config()

# Update configuration
config.update_config('ollama.temperature', 0.2)

# Validate configuration
try:
    config.validate_config()
    print("Configuration is valid")
except ValueError as e:
    print(f"Configuration error: {e}")

# Test Ollama connection
connection_result = config.test_ollama_connection()
if connection_result['success']:
    print("Ollama connection successful")
else:
    print(f"Connection failed: {connection_result['error']}")
```

### Template Validation API

```python
from src.config.template_validator import TemplateValidator

# Load and validate templates
validator = TemplateValidator()
validation_result = validator.validate_all_templates()

if validation_result['valid']:
    print(f"All {validation_result['total_operations']} operations are valid")
else:
    print("Validation errors:")
    for error in validation_result['errors']:
        print(f"  - {error}")

# Get operation summary
summary = validator.get_operation_summary()
for category, info in summary.items():
    print(f"{category}: {info['count']} operations")

# Check specific operation
if validator.validate_operation_exists('data_operations', 'insert_row'):
    config = validator.get_operation_config('data_operations', 'insert_row')
    print(f"Safety level: {config['safety_level']}")
```

### Operation Function Testing

```python
# Test that all operation wrapper functions are available
from operations.chart_operations import shift_axis, transform_values, resize_chart
from operations.query_operations import filter_data, aggregate_data, sort_data
from operations.visualization_operations import create_chart, get_chart_recommendations
from operations.crud_handlers import insert_row, insert_column, update_cells, delete_rows

print("All operation functions are available and ready for template system use")
```

### Comprehensive System Test

```python
# Run the comprehensive test suite
import subprocess
result = subprocess.run(['python', 'test_config_simple.py'], capture_output=True, text=True)
print(result.stdout)
```