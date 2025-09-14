# Operations Module

This module provides comprehensive data manipulation and visualization capabilities for Excel files, including CRUD operations, chart generation, and chart manipulation with built-in safety checks, validation, and confirmation mechanisms.

## Overview

The operations module consists of six main handler classes that work together to provide complete Excel manipulation capabilities:

### Data Operations (CRUD)
1. **DataInsertionHandler** - Create operations (inserting data)
2. **DataQueryHandler** - Read operations (querying and retrieving data)
3. **DataUpdateHandler** - Update operations (modifying existing data)
4. **DataDeletionHandler** - Delete operations (removing data with safety checks)

### Visualization Operations
5. **VisualizationOperations** - Chart generation and creation
6. **ChartManipulator** - Chart manipulation and transformation

## Features

### Safety and Validation
- **Automatic backups** before every operation
- **Safety limits** (max 50 rows for batch operations, max 10,000 for queries)
- **Data type validation** and conversion
- **Confirmation prompts** for destructive operations
- **Preview functionality** for deletions

### Data Operations (CRUD)
- **Flexible targeting** (by row number, range, conditions, unique identifiers)
- **Complex filtering** with comparison operators (=, !=, >, >=, <, <=, contains, starts_with, ends_with)
- **Sorting and aggregation** (sum, avg, min, max, count)
- **Cross-sheet references** support
- **Batch operations** for efficiency

### Visualization Operations
- **Automatic chart type detection** based on data characteristics
- **Multiple chart types** (bar, line, pie, scatter, area, doughnut, radar)
- **Chart positioning and resizing** with precise control
- **Data transformation** operations (add, subtract, multiply, divide)
- **Chart property modification** (title, style, axis labels)
- **Native Excel integration** with embedded chart objects

### Error Handling
- **Comprehensive validation** of all inputs
- **Graceful error recovery** with detailed error messages
- **Operation logging** for audit trails
- **Backup restoration** capabilities

## Quick Start

```python
from src.excel.excel_service import ExcelService
from src.safety.safety_manager import SafetyManager
from src.operations import (
    DataInsertionHandler, DataQueryHandler, 
    DataUpdateHandler, DataDeletionHandler,
    VisualizationOperations, ChartManipulator,
    InsertionData, QueryData, UpdateData, DeletionData
)

# Initialize services
excel_service = ExcelService()
safety_manager = SafetyManager()
excel_service.load_workbook("data.xlsx")

# Initialize CRUD handlers
insert_handler = DataInsertionHandler(excel_service, safety_manager)
query_handler = DataQueryHandler(excel_service, safety_manager)
update_handler = DataUpdateHandler(excel_service, safety_manager)
delete_handler = DataDeletionHandler(excel_service, safety_manager)

# Initialize visualization handlers
viz_operations = VisualizationOperations(excel_service, safety_manager)
chart_manipulator = ChartManipulator(excel_service, safety_manager)
```

## Usage Examples

### Create Operations (Insert)

#### Insert a New Row
```python
# Insert new employee
result = insert_handler.insert_row(InsertionData(
    values=["John Doe", 30, "Engineering", 75000],
    target_sheet="Employees"
))

if result.success:
    print(f"Inserted row: {result.message}")
```

#### Insert a New Column
```python
# Add bonus column
result = insert_handler.insert_column(InsertionData(
    values=[5000, 4500, 6000, 5500],
    target_sheet="Employees",
    column_names=["Bonus"],
    insert_type="column"
))
```

#### Batch Insert Multiple Rows
```python
# Insert multiple employees
rows_data = [
    InsertionData(values=["Alice Smith", 28, "Marketing", 68000], target_sheet="Employees"),
    InsertionData(values=["Bob Johnson", 32, "Sales", 72000], target_sheet="Employees")
]

result = insert_handler.add_multiple_rows(rows_data)
```

### Read Operations (Query)

#### Basic Query
```python
# Get all employees
result = query_handler.query_data(QueryData(
    target_sheet="Employees",
    limit=100
))

for employee in result.data:
    print(f"{employee['Name']} - {employee['Department']}")
```

#### Query with Filters
```python
# Get Engineering employees with salary > 70000
result = query_handler.query_data(QueryData(
    target_sheet="Employees",
    conditions={
        "Department": "Engineering",
        "Salary": {"operator": ">", "value": 70000}
    },
    sort_by="Salary",
    sort_order="desc",
    limit=100
))
```

#### Query with Aggregations
```python
# Get salary statistics
result = query_handler.query_data(QueryData(
    target_sheet="Employees",
    aggregations={
        "Salary": "avg",
        "Age": "avg", 
        "Name": "count"
    },
    limit=100
))

print(f"Average salary: ${result.aggregation_results['avg(Salary)']:.2f}")
```

#### Search Records
```python
# Find employees with "John" in name
result = query_handler.find_records("Employees", "John", ["Name"])
```

### Update Operations

#### Update Single Cell
```python
# Update John's salary
result = update_handler.update_cell("Employees", "D2", 80000)
```

#### Update by Conditions
```python
# Give raises to Engineering department
result = update_handler.update_data(UpdateData(
    target_sheet="Employees",
    conditions={"Department": "Engineering"},
    updates={"Salary": 82000}
))
```

#### Update Specific Row
```python
# Update employee in row 3
result = update_handler.update_data(UpdateData(
    target_sheet="Employees",
    target_row=3,
    updates={"Age": 31, "Salary": 78000}
))
```

#### Update Range
```python
# Update salary range
result = update_handler.update_range(
    "Employees", 
    "D2:D4", 
    [[75000], [78000], [82000]]
)
```

### Delete Operations

#### Delete with Confirmation
```python
# Delete Marketing employees (requires confirmation)
result = delete_handler.delete_by_condition(
    "Employees", 
    {"Department": "Marketing"}
)

if result.requires_confirmation:
    print(result.confirmation_prompt)
    # User confirms...
    result = delete_handler.delete_by_condition(
        "Employees", 
        {"Department": "Marketing"},
        confirmed=True
    )
```

#### Delete Specific Rows
```python
# Delete rows 5 and 7
result = delete_handler.delete_rows("Employees", [5, 7], confirmed=True)
```

#### Clear Range
```python
# Clear bonus data
result = delete_handler.delete_data(DeletionData(
    target_sheet="Employees",
    target_range="E2:E10",
    confirmation_required=False
), confirmed=True)
```

## Visualization Operations

### Chart Creation

#### Automatic Chart Type Detection
```python
# Get chart recommendations based on data
recommendations = viz_operations.get_chart_recommendations(
    sheet_name="Sales",
    data_range="A1:C10"
)

print(f"Recommended chart type: {recommendations.recommended_type}")
print(f"Confidence: {recommendations.confidence}")
```

#### Create Charts with Auto-Detection
```python
# Create chart with automatic type detection
result = viz_operations.create_chart(
    sheet_name="Sales",
    data_range="A1:C10",
    title="Monthly Sales Report"
)

if result.success:
    print(f"Created {result.chart_type} chart: {result.chart_id}")
```

#### Create Specific Chart Types
```python
# Create bar chart
result = viz_operations.create_chart(
    sheet_name="Sales",
    data_range="A1:C10",
    chart_type="bar",
    title="Sales by Region",
    x_axis_title="Region",
    y_axis_title="Sales ($)"
)

# Create line chart for time series
result = viz_operations.create_chart(
    sheet_name="Trends",
    data_range="A1:B20",
    chart_type="line",
    title="Monthly Revenue Trend",
    width=600,
    height=400
)

# Create pie chart for categorical data
result = viz_operations.create_chart(
    sheet_name="Categories",
    data_range="A1:B6",
    chart_type="pie",
    title="Market Share Distribution"
)
```

### Chart Manipulation

#### Chart Positioning
```python
# Shift chart along X-axis (horizontal)
result = chart_manipulator.shift_chart_axis(
    chart_id="Chart1",
    axis="x",
    amount=-2  # Move left by 2 units
)

# Shift chart along Y-axis (vertical)
result = chart_manipulator.shift_chart_axis(
    chart_id="Chart1", 
    axis="y",
    amount=3   # Move down by 3 units
)
```

#### Chart Resizing
```python
# Resize chart dimensions
result = chart_manipulator.resize_chart(
    chart_id="Chart1",
    width=800,
    height=600
)
```

#### Data Transformation
```python
# Subtract value from all Y-axis data
result = chart_manipulator.transform_chart_values(
    chart_id="Chart1",
    axis="y",
    operation="subtract",
    value=1000
)

# Multiply all chart values
result = chart_manipulator.transform_chart_values(
    chart_id="Chart1",
    axis="y", 
    operation="multiply",
    value=1.5
)
```

#### Chart Property Modification
```python
# Modify chart properties
result = chart_manipulator.modify_chart_properties(
    chart_id="Chart1",
    property="title",
    value="Updated Sales Report"
)

# Update axis titles
result = chart_manipulator.modify_chart_properties(
    chart_id="Chart1",
    property="x_axis_title",
    value="Time Period"
)
```

#### Chart Management
```python
# List all charts in workbook
charts = chart_manipulator.list_charts()
for chart in charts:
    print(f"Chart ID: {chart.chart_id}, Title: {chart.title}")

# Get chart information
chart_info = chart_manipulator.get_chart_info("Chart1")
print(f"Chart type: {chart_info.chart_type}")
print(f"Position: ({chart_info.x}, {chart_info.y})")
print(f"Size: {chart_info.width}x{chart_info.height}")
```

## Data Structures

### InsertionData
```python
@dataclass
class InsertionData:
    values: List[Any]                    # Values to insert
    target_sheet: str                    # Target sheet name
    target_row: Optional[int] = None     # Specific row (for row insertion)
    target_column: Optional[str] = None  # Specific column (for column insertion)
    column_names: Optional[List[str]] = None  # Column headers
    insert_type: str = "row"             # "row" or "column"
```

### QueryData
```python
@dataclass
class QueryData:
    target_sheet: str                           # Target sheet name
    columns: Optional[List[str]] = None         # Columns to select
    conditions: Optional[Dict[str, Any]] = None # Filter conditions
    sort_by: Optional[str] = None               # Sort column
    sort_order: str = "asc"                     # "asc" or "desc"
    limit: Optional[int] = None                 # Result limit
    aggregations: Optional[Dict[str, str]] = None  # Aggregation functions
```

### UpdateData
```python
@dataclass
class UpdateData:
    target_sheet: str                              # Target sheet name
    updates: Dict[str, Any]                        # Column -> value updates
    conditions: Optional[Dict[str, Any]] = None    # Filter conditions
    target_row: Optional[int] = None               # Specific row
    target_range: Optional[str] = None             # Specific range
    unique_identifier: Optional[Dict[str, Any]] = None  # Unique identifier
```

### DeletionData
```python
@dataclass
class DeletionData:
    target_sheet: str                              # Target sheet name
    conditions: Optional[Dict[str, Any]] = None    # Filter conditions
    target_rows: Optional[List[int]] = None        # Specific rows
    target_range: Optional[str] = None             # Specific range
    unique_identifier: Optional[Dict[str, Any]] = None  # Unique identifier
    confirmation_required: bool = True             # Require confirmation
```

### Visualization Data Structures

#### ChartConfig
```python
@dataclass
class ChartConfig:
    chart_type: str                                # Chart type (bar, line, pie, etc.)
    title: Optional[str] = None                    # Chart title
    x_axis_title: Optional[str] = None             # X-axis label
    y_axis_title: Optional[str] = None             # Y-axis label
    width: int = 400                               # Chart width in pixels
    height: int = 300                              # Chart height in pixels
    x: int = 100                                   # X position in worksheet
    y: int = 50                                    # Y position in worksheet
    style: Optional[str] = None                    # Chart style
```

#### ChartRecommendation
```python
@dataclass
class ChartRecommendation:
    recommended_type: str                          # Recommended chart type
    confidence: float                              # Confidence score (0-1)
    reasoning: str                                 # Explanation for recommendation
    alternatives: List[str]                        # Alternative chart types
    data_characteristics: Dict[str, Any]           # Analysis of data properties
```

#### ChartInfo
```python
@dataclass
class ChartInfo:
    chart_id: str                                  # Unique chart identifier
    chart_type: str                                # Chart type
    title: str                                     # Chart title
    sheet_name: str                                # Sheet containing chart
    data_range: str                                # Source data range
    x: int                                         # X position
    y: int                                         # Y position
    width: int                                     # Chart width
    height: int                                    # Chart height
```

## Condition Operators

The following operators are supported in filter conditions:

- `=` - Equality (default)
- `!=` - Not equal
- `>` - Greater than
- `>=` - Greater than or equal
- `<` - Less than
- `<=` - Less than or equal
- `contains` - String contains (case-insensitive)
- `starts_with` - String starts with (case-insensitive)
- `ends_with` - String ends with (case-insensitive)

### Example Conditions
```python
conditions = {
    "Age": {"operator": ">=", "value": 30},
    "Department": "Engineering",
    "Name": {"operator": "contains", "value": "John"},
    "Salary": {"operator": ">", "value": 70000}
}
```

## Safety Features

### Automatic Backups
- Created before every operation (not just destructive ones)
- Timestamped backup files with retention policy
- Configurable backup directory and retention count

### Operation Limits
- **Row operations**: Maximum 50 rows per batch operation
- **Query results**: Maximum 10,000 rows per query
- **Column insertion**: Maximum 1,000 values per column
- **Row insertion**: Maximum 100 values per row

### Confirmation System
- **Preview functionality** shows what will be affected
- **Confirmation prompts** with sample data
- **Safety warnings** for potentially dangerous operations
- **Unique identifier validation** warns about ambiguous matches

### Data Validation
- **Type conversion** based on column data types
- **Range validation** for cell references
- **Sheet existence** checks
- **Parameter validation** for all operations

## Error Handling

All operations return result objects with:
- `success`: Boolean indicating operation success
- `message`: Descriptive message about the operation
- `affected_rows/affected_cells`: Count of affected data
- `warnings`: List of warning messages
- `data`: Operation-specific result data

### Example Error Handling
```python
result = query_handler.query_data(query_data)

if result.success:
    print(f"Query successful: {result.message}")
    print(f"Retrieved {result.row_count} rows")
    
    if result.warnings:
        for warning in result.warnings:
            print(f"Warning: {warning}")
else:
    print(f"Query failed: {result.message}")
```

## Integration with Other Modules

The CRUD operations module integrates with:

- **ExcelService**: For file operations and structure analysis
- **SafetyManager**: For operation validation and risk assessment
- **Template System**: For dynamic operation definitions
- **LLM Service**: For natural language command interpretation

## Testing

Comprehensive test suites are provided:

### CRUD Operations Tests
- `tests/test_crud_insertion.py` - Tests for create operations
- `tests/test_crud_query.py` - Tests for read operations  
- `tests/test_crud_update.py` - Tests for update operations
- `tests/test_crud_deletion.py` - Tests for delete operations

### Visualization Operations Tests
- `tests/test_visualization_operations.py` - Tests for chart generation
- `tests/test_chart_operations.py` - Tests for chart manipulation

Run tests with:
```bash
# All operations tests
python -m pytest tests/test_*operations*.py -v

# CRUD operations only
python -m pytest tests/test_crud_*.py -v

# Visualization operations only
python -m pytest tests/test_*visualization*.py tests/test_*chart*.py -v
```

## Demos

Comprehensive demonstration scripts are provided:

### CRUD Operations Demo
```bash
python examples/crud_operations_demo.py
```
Demonstrates all CRUD operations working together with sample data.

### Visualization Demo
```bash
python examples/visualization_demo.py
```
Shows chart generation, manipulation, and various chart types.

### Simple Test Scripts
```bash
# Simple visualization test
python test_visualization_simple.py

# Chart manipulation test
python test_chart_manipulation.py
```

## Requirements Fulfilled

This implementation fulfills the following requirements from the specification:

### Requirement 1 (Create Operations)
- ✅ 1.1: Parse natural language commands for data creation
- ✅ 1.2: Validate data types and column structure
- ✅ 1.3: Provide confirmation of successful operations
- ✅ 1.4: Handle missing locations with clarification

### Requirement 2 (Read Operations)  
- ✅ 2.1: Interpret queries and return relevant results
- ✅ 2.2: Support filtering, sorting, and aggregation
- ✅ 2.3: Present results in clear, readable format
- ✅ 2.4: Ask clarifying questions for unclear queries

### Requirement 3 (Update Operations)
- ✅ 3.1: Locate target cells and apply changes
- ✅ 3.2: Preserve data integrity and validate values
- ✅ 3.3: Confirm changes made
- ✅ 3.4: Request specificity for ambiguous updates

### Requirement 4 (Delete Operations)
- ✅ 4.1: Identify and remove target entries
- ✅ 4.2: Ask for confirmation before deletion
- ✅ 4.3: Report what was removed
- ✅ 4.4: Seek clarification for ambiguous requests

### Requirement 5 (Visualization Operations)
- ✅ 5.1: Analyze data and create appropriate visualizations
- ✅ 5.2: Support common chart types (bar, line, pie, scatter, etc.)
- ✅ 5.3: Embed plots directly in Excel files
- ✅ 5.4: Suggest chart types based on data characteristics
- ✅ 5.5: Explain visualization limitations and suggest alternatives

### Chart Manipulation Requirements
- ✅ Chart positioning and axis shifting
- ✅ Data transformation operations
- ✅ Chart resizing and formatting
- ✅ Chart reference management system

All operations include comprehensive safety checks, validation, and error handling as specified in the design requirements. The visualization system provides intelligent chart type detection and native Excel integration for seamless chart creation and manipulation.