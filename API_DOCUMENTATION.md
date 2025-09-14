# Excel-LLM Integration Tool - API Documentation

## Overview

This document provides comprehensive API documentation for the Excel-LLM Integration Tool, covering all modules, classes, and functions available for developers.

## Table of Contents

1. [Core Services](#core-services)
2. [CRUD Operations](#crud-operations)
3. [Visualization Operations](#visualization-operations)
4. [Safety System](#safety-system)
5. [Template System](#template-system)
6. [LLM Service](#llm-service)
7. [Configuration](#configuration)
8. [Data Models](#data-models)

## Core Services

### ExcelService

Main service for Excel file operations with automatic backup functionality.

```python
from src.excel.excel_service import ExcelService

class ExcelService:
    def __init__(self, backup_dir: str = "./backups", max_backups: int = 10)
    def load_workbook(self, file_path: str) -> bool
    def create_backup(self) -> Optional[str]
    def save_workbook(self, create_backup: bool = True) -> bool
    def get_structure(self) -> Optional[ExcelStructure]
    def get_sheet_names(self) -> List[str]
    def get_sheet(self, sheet_name: str) -> Optional[Worksheet]
    def close(self) -> None
```

**Key Methods:**
- `load_workbook()`: Load Excel file and analyze structure
- `create_backup()`: Create timestamped backup with retention policy
- `get_structure()`: Get comprehensive workbook analysis
- `save_workbook()`: Save with optional automatic backup

### SafetyManager

Coordinates all safety mechanisms for Excel operations.

```python
from src.safety.safety_manager import SafetyManager

class SafetyManager:
    def __init__(self, max_rows: int = 50, max_columns: int = 20)
    def evaluate_operation(self, operation: str, parameters: Dict[str, Any], 
                          command_text: str = "") -> SafetyResult
    def is_operation_safe(self, operation: str, parameters: Dict[str, Any]) -> bool
    def create_confirmation_prompt(self, safety_result: SafetyResult) -> str
    def get_safety_summary(self, safety_result: SafetyResult) -> str
```

**Key Methods:**
- `evaluate_operation()`: Comprehensive safety analysis
- `is_operation_safe()`: Quick safety check
- `create_confirmation_prompt()`: Generate user confirmation prompts

## CRUD Operations

### DataInsertionHandler

Handles data creation operations with validation and safety checks.

```python
from src.operations.crud_handlers import DataInsertionHandler, InsertionData

class DataInsertionHandler:
    def __init__(self, excel_service: ExcelService, safety_manager: SafetyManager)
    def insert_row(self, data: InsertionData) -> OperationResult
    def insert_column(self, data: InsertionData) -> OperationResult
    def add_multiple_rows(self, rows_data: List[InsertionData]) -> OperationResult
```

**Usage Example:**
```python
result = handler.insert_row(InsertionData(
    values=["John Doe", 30, "Engineering"],
    target_sheet="Employees"
))
```

### DataQueryHandler

Handles data retrieval with filtering, sorting, and aggregation.

```python
class DataQueryHandler:
    def __init__(self, excel_service: ExcelService, safety_manager: SafetyManager)
    def query_data(self, query: QueryData) -> QueryResult
    def find_records(self, sheet_name: str, search_term: str, 
                    columns: List[str] = None) -> QueryResult
    def get_unique_values(self, sheet_name: str, column: str) -> List[Any]
```

**Usage Example:**
```python
result = handler.query_data(QueryData(
    target_sheet="Sales",
    conditions={"Region": "North", "Amount": {"operator": ">", "value": 1000}},
    sort_by="Amount",
    sort_order="desc"
))
```

### DataUpdateHandler

Handles data modification with integrity checks.

```python
class DataUpdateHandler:
    def __init__(self, excel_service: ExcelService, safety_manager: SafetyManager)
    def update_data(self, update: UpdateData) -> OperationResult
    def update_cell(self, sheet_name: str, cell_ref: str, value: Any) -> OperationResult
    def update_range(self, sheet_name: str, range_ref: str, 
                    values: List[List[Any]]) -> OperationResult
```

**Usage Example:**
```python
result = handler.update_data(UpdateData(
    target_sheet="Employees",
    conditions={"Department": "Engineering"},
    updates={"Salary": 85000}
))
```

### DataDeletionHandler

Handles data removal with confirmation and safety limits.

```python
class DataDeletionHandler:
    def __init__(self, excel_service: ExcelService, safety_manager: SafetyManager)
    def delete_data(self, deletion: DeletionData, confirmed: bool = False) -> OperationResult
    def delete_rows(self, sheet_name: str, row_numbers: List[int], 
                   confirmed: bool = False) -> OperationResult
    def delete_by_condition(self, sheet_name: str, conditions: Dict[str, Any], 
                           confirmed: bool = False) -> OperationResult
```

**Usage Example:**
```python
result = handler.delete_by_condition(
    "Employees",
    {"Status": "Inactive"},
    confirmed=True
)
```

## Visualization Operations

### VisualizationOperations

Main class for chart generation with automatic type detection.

```python
from src.operations.visualization_operations import VisualizationOperations

class VisualizationOperations:
    def __init__(self, excel_service: ExcelService, safety_manager: SafetyManager)
    def create_chart(self, sheet_name: str, data_range: str, 
                    chart_type: str = None, **kwargs) -> OperationResult
    def get_chart_recommendations(self, sheet_name: str, 
                                 data_range: str) -> ChartRecommendation
```

**Supported Chart Types:**
- `bar` - Bar charts for categorical comparisons
- `line` - Line charts for time series data
- `pie` - Pie charts for proportional data
- `scatter` - Scatter plots for correlation analysis
- `area` - Area charts for cumulative data
- `doughnut` - Doughnut charts (pie chart variant)
- `radar` - Radar charts for multi-dimensional data

**Usage Example:**
```python
# Automatic chart type detection
result = viz_ops.create_chart(
    sheet_name="Sales",
    data_range="A1:C10",
    title="Monthly Sales Report"
)

# Specific chart type
result = viz_ops.create_chart(
    sheet_name="Sales",
    data_range="A1:C10",
    chart_type="bar",
    title="Sales by Region",
    x_axis_title="Region",
    y_axis_title="Sales ($)",
    width=600,
    height=400
)
```

### ChartManipulator

Handles chart positioning, resizing, and data transformation.

```python
from src.operations.chart_operations import ChartManipulator

class ChartManipulator:
    def __init__(self, excel_service: ExcelService, safety_manager: SafetyManager)
    def shift_chart_axis(self, chart_id: str, axis: str, amount: float) -> OperationResult
    def resize_chart(self, chart_id: str, width: int, height: int) -> OperationResult
    def transform_chart_values(self, chart_id: str, axis: str, 
                              operation: str, value: float) -> OperationResult
    def modify_chart_properties(self, chart_id: str, property: str, 
                               value: Any) -> OperationResult
    def list_charts(self) -> List[ChartInfo]
```

**Usage Example:**
```python
# Move chart position
result = chart_ops.shift_chart_axis("Chart1", "x", -2)

# Resize chart
result = chart_ops.resize_chart("Chart1", 800, 600)

# Transform data (subtract 1000 from all Y values)
result = chart_ops.transform_chart_values("Chart1", "y", "subtract", 1000)
```

## Safety System

### RiskAssessor

Classifies operations into risk levels.

```python
from src.safety.risk_assessor import RiskAssessor, RiskLevel

class RiskAssessor:
    def assess_operation(self, operation: str, parameters: Dict[str, Any], 
                        command_text: str = "") -> RiskAssessment
    def get_risk_explanation(self, assessment: RiskAssessment) -> str
```

**Risk Levels:**
- `LOW` - Safe operations (read, simple queries)
- `MEDIUM` - Structural changes (updates, chart modifications)
- `HIGH` - Destructive operations (deletions)
- `DANGEROUS` - Mass operations (blocked automatically)

### ScopeAnalyzer

Analyzes and enforces operation scope limits.

```python
from src.safety.scope_analyzer import ScopeAnalyzer

class ScopeAnalyzer:
    def __init__(self, max_rows: int = 50, max_columns: int = 20)
    def analyze_scope(self, operation: str, parameters: Dict[str, Any], 
                     sheet_info: Optional[Dict[str, Any]] = None) -> ScopeAnalysis
```

### CommandBlocker

Blocks dangerous operations with helpful error messages.

```python
from src.safety.command_blocker import CommandBlocker

class CommandBlocker:
    def check_command(self, operation: str, parameters: Dict[str, Any], 
                     command_text: str = "") -> BlockResult
    def get_blocked_operations(self) -> List[str]
```

### ParameterValidator

Validates and sanitizes operation parameters.

```python
from src.safety.parameter_validator import ParameterValidator

class ParameterValidator:
    def validate_parameters(self, operation: str, 
                          parameters: Dict[str, Any]) -> ValidationResult
    def get_parameter_help(self, operation: str) -> Dict[str, str]
```

## Template System

### TemplateLoader

Loads operation definitions from YAML configuration.

```python
from src.templates.template_loader import TemplateLoader

class TemplateLoader:
    def __init__(self, templates_dir: Optional[str] = None)
    def load_operations(self) -> Dict[str, Any]
    def get_operations_by_category(self, category: str) -> Dict[str, Any]
    def get_operations_by_safety_level(self, safety_level: str) -> Dict[str, Dict[str, Any]]
    def search_operations_by_keyword(self, keyword: str) -> Dict[str, Dict[str, Any]]
```

### TemplateRegistry

Maps operation names to their corresponding functions.

```python
from src.templates.template_registry import TemplateRegistry

class TemplateRegistry:
    def __init__(self, template_loader: Optional[TemplateLoader] = None)
    def get_function(self, category: str, operation: str) -> Optional[Callable]
    def execute_operation(self, category: str, operation: str, *args, **kwargs) -> Any
    def is_operation_available(self, category: str, operation: str) -> bool
    def get_registry_stats(self) -> Dict[str, Any]
```

### PromptGenerator

Generates dynamic LLM prompts based on available operations.

```python
from src.templates.prompt_generator import PromptGenerator

class PromptGenerator:
    def __init__(self, template_loader: Optional[TemplateLoader] = None,
                 template_registry: Optional[TemplateRegistry] = None)
    def generate_system_prompt(self) -> str
    def generate_operation_summary(self) -> str
    def suggest_operations_for_command(self, user_command: str) -> List[Dict[str, Any]]
```

### HotReloadManager

Manages hot-reload functionality for operation configurations.

```python
from src.templates.hot_reload import HotReloadManager

class HotReloadManager:
    def __init__(self, template_loader: Optional[TemplateLoader] = None, ...)
    def start_watching(self, watch_directory: Optional[str] = None) -> bool
    def stop_watching(self) -> bool
    def manual_reload(self) -> bool
    def add_reload_callback(self, name: str, callback: Callable[[], None]) -> None
```

## LLM Service

### OllamaService

Service for interacting with Ollama LLM for command processing.

```python
from src.llm.ollama_service import OllamaService, LLMResponse

class OllamaService:
    def __init__(self)
    def initialize_connection(self) -> bool
    def parse_to_structured_command(self, user_command: str) -> LLMResponse
    def assess_command_safety(self, command: str) -> str
    def generate_confirmation_prompt(self, operation: Dict[str, Any]) -> str
    def validate_response(self, response: LLMResponse) -> bool
```

**Usage Example:**
```python
llm_service = OllamaService()
llm_service.initialize_connection()

response = llm_service.parse_to_structured_command(
    "Create a bar chart from sales data in columns A to C"
)

print(f"Intent: {response.intent}")
print(f"Operation: {response.operation}")
print(f"Parameters: {response.parameters}")
```

## Configuration

### ConfigManager

Manages application configuration from YAML files.

```python
from src.config.config_manager import ConfigManager, config

class ConfigManager:
    def __init__(self, config_path: Optional[str] = None)
    def get(self, key: str, default: Any = None) -> Any
    def get_ollama_config(self) -> Dict[str, Any]
    def get_backup_config(self) -> Dict[str, Any]
    def get_safety_config(self) -> Dict[str, Any]
    def validate_config(self) -> bool

# Global configuration instance
config = ConfigManager()
```

## Data Models

### Core Data Structures

#### InsertionData
```python
@dataclass
class InsertionData:
    values: List[Any]
    target_sheet: str
    target_row: Optional[int] = None
    target_column: Optional[str] = None
    column_names: Optional[List[str]] = None
    insert_type: str = "row"
```

#### QueryData
```python
@dataclass
class QueryData:
    target_sheet: str
    columns: Optional[List[str]] = None
    conditions: Optional[Dict[str, Any]] = None
    sort_by: Optional[str] = None
    sort_order: str = "asc"
    limit: Optional[int] = None
    aggregations: Optional[Dict[str, str]] = None
```

#### UpdateData
```python
@dataclass
class UpdateData:
    target_sheet: str
    updates: Dict[str, Any]
    conditions: Optional[Dict[str, Any]] = None
    target_row: Optional[int] = None
    target_range: Optional[str] = None
    unique_identifier: Optional[Dict[str, Any]] = None
```

#### DeletionData
```python
@dataclass
class DeletionData:
    target_sheet: str
    conditions: Optional[Dict[str, Any]] = None
    target_rows: Optional[List[int]] = None
    target_range: Optional[str] = None
    unique_identifier: Optional[Dict[str, Any]] = None
    confirmation_required: bool = True
```

### Visualization Data Structures

#### ChartConfig
```python
@dataclass
class ChartConfig:
    chart_type: str
    title: Optional[str] = None
    x_axis_title: Optional[str] = None
    y_axis_title: Optional[str] = None
    width: int = 400
    height: int = 300
    x: int = 100
    y: int = 50
    style: Optional[str] = None
```

#### ChartRecommendation
```python
@dataclass
class ChartRecommendation:
    recommended_type: str
    confidence: float
    reasoning: str
    alternatives: List[str]
    data_characteristics: Dict[str, Any]
```

#### ChartInfo
```python
@dataclass
class ChartInfo:
    chart_id: str
    chart_type: str
    title: str
    sheet_name: str
    data_range: str
    x: int
    y: int
    width: int
    height: int
```

### Safety Data Structures

#### SafetyResult
```python
@dataclass
class SafetyResult:
    safe: bool
    blocked: bool
    confirmation_required: bool
    risk_assessment: RiskAssessment
    scope_analysis: ScopeAnalysis
    block_result: BlockResult
    validation_result: ValidationResult
    error_messages: List[str]
    warnings: List[str]
    suggestions: List[str]
```

#### OperationResult
```python
@dataclass
class OperationResult:
    success: bool
    message: str
    affected_rows: int = 0
    affected_cells: int = 0
    warnings: List[str] = field(default_factory=list)
    data: Optional[Any] = None
```

### LLM Data Structures

#### LLMResponse
```python
@dataclass
class LLMResponse:
    intent: str
    operation: str
    parameters: Dict[str, Any]
    confirmation_required: bool
    risk_assessment: str
    confidence: float = 0.0
    raw_response: str = ""
```

## Error Handling

All API methods follow consistent error handling patterns:

1. **Return Values**: Operations return result objects with success indicators
2. **Exceptions**: Critical errors raise specific exception types
3. **Validation**: Input validation with detailed error messages
4. **Logging**: Comprehensive logging for debugging and audit trails

### Common Exception Types

```python
# LLM Service
class OllamaConnectionError(Exception): pass

# Excel Service  
class ExcelServiceError(Exception): pass

# Safety System
class SafetyViolationError(Exception): pass

# Template System
class TemplateLoadError(Exception): pass
```

### Error Response Format

```python
{
    "success": False,
    "message": "Detailed error description",
    "error_code": "ERROR_TYPE",
    "suggestions": ["Alternative action 1", "Alternative action 2"]
}
```

## Integration Examples

### Complete Workflow Example

```python
from src.excel.excel_service import ExcelService
from src.safety.safety_manager import SafetyManager
from src.operations.crud_handlers import DataQueryHandler
from src.operations.visualization_operations import VisualizationOperations
from src.llm.ollama_service import OllamaService

# Initialize services
excel_service = ExcelService()
safety_manager = SafetyManager()
llm_service = OllamaService()

# Load Excel file
excel_service.load_workbook("sales_data.xlsx")

# Parse natural language command
response = llm_service.parse_to_structured_command(
    "Create a bar chart showing sales by region"
)

# Execute visualization operation
if response.intent == "visualization_operations":
    viz_ops = VisualizationOperations(excel_service, safety_manager)
    result = viz_ops.create_chart(**response.parameters)
    
    if result.success:
        print(f"Chart created successfully: {result.message}")
    else:
        print(f"Chart creation failed: {result.message}")

# Save with backup
excel_service.save_workbook(create_backup=True)
```

This API documentation provides comprehensive coverage of all available functionality in the Excel-LLM Integration Tool. For additional examples and usage patterns, refer to the demo scripts in the `examples/` directory.