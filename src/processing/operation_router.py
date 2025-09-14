"""
Operation router that routes commands to appropriate handlers based on structured LLM output.

This module provides intelligent routing of operations to the correct handlers
and manages the execution flow between different operation types.
"""

import logging
from typing import Dict, Any, Optional, List, Callable
from dataclasses import dataclass
from enum import Enum

from ..llm.ollama_service import LLMResponse
from ..templates.template_registry import TemplateRegistry
from ..operations.crud_handlers import (
    DataInsertionHandler, DataQueryHandler, 
    InsertionData, QueryData, OperationResult
)
from operations.visualization_operations import VisualizationOperations
from operations.chart_operations import ChartManipulator
from ..excel.excel_service import ExcelService


class RoutingStrategy(Enum):
    """Strategy for routing operations."""
    TEMPLATE_REGISTRY = "template_registry"
    DIRECT_HANDLER = "direct_handler"
    HYBRID = "hybrid"


@dataclass
class RoutingResult:
    """Result of operation routing."""
    success: bool
    handler_used: str
    execution_result: Any
    error_message: Optional[str] = None
    warnings: Optional[List[str]] = None


class OperationRouter:
    """
    Routes operations to appropriate handlers based on structured LLM output.
    
    This class determines the best execution path for each operation and manages
    the coordination between different operation handlers.
    """
    
    def __init__(self, 
                 template_registry: TemplateRegistry,
                 excel_service: ExcelService,
                 routing_strategy: RoutingStrategy = RoutingStrategy.HYBRID):
        """
        Initialize operation router.
        
        Args:
            template_registry: Template registry for operation execution
            excel_service: Excel service for file operations
            routing_strategy: Strategy for routing operations
        """
        self.template_registry = template_registry
        self.excel_service = excel_service
        self.routing_strategy = routing_strategy
        self.logger = logging.getLogger(__name__)
        
        # Initialize operation handlers
        self._initialize_handlers()
        
        # Define routing rules
        self._routing_rules = self._build_routing_rules()
    
    def _initialize_handlers(self):
        """Initialize all operation handlers."""
        # Import safety manager (assuming it's available)
        from ..safety.safety_manager import SafetyManager
        safety_manager = SafetyManager()
        
        # Initialize handlers
        self.data_insertion_handler = DataInsertionHandler(self.excel_service, safety_manager)
        self.data_query_handler = DataQueryHandler(self.excel_service, safety_manager)
        self.visualization_operations = VisualizationOperations()
        
        # Initialize chart operations if available
        try:
            self.chart_operations = ChartManipulator()
        except Exception as e:
            self.logger.warning(f"Chart operations not available: {e}")
            self.chart_operations = None
    
    def _build_routing_rules(self) -> Dict[str, Dict[str, Callable]]:
        """Build routing rules for different operation types."""
        return {
            "data_operations": {
                "create_data": self._route_data_creation,
                "insert_row": self._route_data_creation,
                "insert_column": self._route_data_creation,
                "query_data": self._route_data_query,
                "find_records": self._route_data_query,
                "get_summary": self._route_data_query,
                "update_data": self._route_data_update,
                "delete_data": self._route_data_deletion
            },
            "visualization_operations": {
                "create_chart": self._route_chart_creation,
                "generate_plot": self._route_chart_creation,
                "visualize_data": self._route_chart_creation
            },
            "chart_operations": {
                "shift_axis": self._route_chart_manipulation,
                "transform_values": self._route_chart_manipulation,
                "resize_chart": self._route_chart_manipulation,
                "move_chart": self._route_chart_manipulation
            }
        }
    
    def route_operation(self, llm_response: LLMResponse) -> RoutingResult:
        """
        Route operation to appropriate handler.
        
        Args:
            llm_response: Structured LLM response with operation details
            
        Returns:
            RoutingResult: Result of routing and execution
        """
        intent = llm_response.intent
        operation = llm_response.operation
        parameters = llm_response.parameters
        
        self.logger.info(f"Routing operation: {intent}.{operation}")
        
        try:
            # Check if we have a specific routing rule
            if intent in self._routing_rules and operation in self._routing_rules[intent]:
                router_func = self._routing_rules[intent][operation]
                return router_func(llm_response)
            
            # Fallback to strategy-based routing
            return self._strategy_based_routing(llm_response)
            
        except Exception as e:
            self.logger.error(f"Error routing operation {intent}.{operation}: {str(e)}")
            return RoutingResult(
                success=False,
                handler_used="error",
                execution_result=None,
                error_message=f"Routing error: {str(e)}"
            )
    
    def _strategy_based_routing(self, llm_response: LLMResponse) -> RoutingResult:
        """Route operation based on configured strategy."""
        if self.routing_strategy == RoutingStrategy.TEMPLATE_REGISTRY:
            return self._route_via_template_registry(llm_response)
        elif self.routing_strategy == RoutingStrategy.DIRECT_HANDLER:
            return self._route_via_direct_handler(llm_response)
        else:  # HYBRID
            return self._route_via_hybrid_approach(llm_response)
    
    def _route_via_template_registry(self, llm_response: LLMResponse) -> RoutingResult:
        """Route operation through template registry."""
        try:
            if self.template_registry.is_operation_available(llm_response.intent, llm_response.operation):
                result = self.template_registry.execute_operation(
                    llm_response.intent, 
                    llm_response.operation, 
                    **llm_response.parameters
                )
                
                return RoutingResult(
                    success=True,
                    handler_used="template_registry",
                    execution_result=result
                )
            else:
                return RoutingResult(
                    success=False,
                    handler_used="template_registry",
                    execution_result=None,
                    error_message=f"Operation {llm_response.intent}.{llm_response.operation} not available in template registry"
                )
                
        except NotImplementedError as e:
            return RoutingResult(
                success=False,
                handler_used="template_registry",
                execution_result=None,
                error_message=f"Operation not implemented: {str(e)}"
            )
        except Exception as e:
            return RoutingResult(
                success=False,
                handler_used="template_registry",
                execution_result=None,
                error_message=f"Template registry execution error: {str(e)}"
            )
    
    def _route_via_direct_handler(self, llm_response: LLMResponse) -> RoutingResult:
        """Route operation directly to appropriate handler."""
        intent = llm_response.intent
        operation = llm_response.operation
        
        # Try to find a direct handler
        if intent in self._routing_rules and operation in self._routing_rules[intent]:
            router_func = self._routing_rules[intent][operation]
            return router_func(llm_response)
        
        return RoutingResult(
            success=False,
            handler_used="direct_handler",
            execution_result=None,
            error_message=f"No direct handler found for {intent}.{operation}"
        )
    
    def _route_via_hybrid_approach(self, llm_response: LLMResponse) -> RoutingResult:
        """Route operation using hybrid approach (template registry first, then direct)."""
        # Try template registry first
        template_result = self._route_via_template_registry(llm_response)
        
        if template_result.success:
            return template_result
        
        # If template registry fails, try direct handler
        self.logger.info(f"Template registry failed for {llm_response.intent}.{llm_response.operation}, trying direct handler")
        direct_result = self._route_via_direct_handler(llm_response)
        
        if direct_result.success:
            # Add warning about fallback
            if direct_result.warnings is None:
                direct_result.warnings = []
            direct_result.warnings.append("Operation executed via direct handler (template registry unavailable)")
            return direct_result
        
        # Both failed
        return RoutingResult(
            success=False,
            handler_used="hybrid",
            execution_result=None,
            error_message=f"Both template registry and direct handler failed for {llm_response.intent}.{llm_response.operation}",
            warnings=[template_result.error_message, direct_result.error_message]
        )
    
    # Specific routing methods for different operation types
    
    def _route_data_creation(self, llm_response: LLMResponse) -> RoutingResult:
        """Route data creation operations."""
        try:
            parameters = llm_response.parameters
            
            # Create InsertionData object
            insertion_data = InsertionData(
                values=parameters.get('data', parameters.get('values', [])),
                target_sheet=parameters.get('sheet_name', parameters.get('target_sheet', '')),
                target_row=parameters.get('target_row'),
                target_column=parameters.get('target_column'),
                column_names=parameters.get('column_names'),
                insert_type=parameters.get('insert_type', 'row')
            )
            
            # Execute based on insert type
            if insertion_data.insert_type == 'column':
                result = self.data_insertion_handler.insert_column(insertion_data)
            else:
                result = self.data_insertion_handler.insert_row(insertion_data)
            
            return RoutingResult(
                success=result.success,
                handler_used="data_insertion_handler",
                execution_result=result
            )
            
        except Exception as e:
            return RoutingResult(
                success=False,
                handler_used="data_insertion_handler",
                execution_result=None,
                error_message=f"Data creation error: {str(e)}"
            )
    
    def _route_data_query(self, llm_response: LLMResponse) -> RoutingResult:
        """Route data query operations."""
        try:
            parameters = llm_response.parameters
            
            # Handle different query types
            if llm_response.operation == "find_records":
                result = self.data_query_handler.find_records(
                    sheet_name=parameters.get('sheet_name', ''),
                    search_term=parameters.get('search_term', ''),
                    columns=parameters.get('columns')
                )
            elif llm_response.operation == "get_summary":
                result = self.data_query_handler.get_sheet_summary(
                    sheet_name=parameters.get('sheet_name', '')
                )
            else:
                # Standard query
                query_data = QueryData(
                    target_sheet=parameters.get('sheet_name', parameters.get('target_sheet', '')),
                    columns=parameters.get('columns'),
                    conditions=parameters.get('conditions'),
                    sort_by=parameters.get('sort_by'),
                    sort_order=parameters.get('sort_order', 'asc'),
                    limit=parameters.get('limit', 100),
                    aggregations=parameters.get('aggregations')
                )
                result = self.data_query_handler.query_data(query_data)
            
            return RoutingResult(
                success=result.success,
                handler_used="data_query_handler",
                execution_result=result
            )
            
        except Exception as e:
            return RoutingResult(
                success=False,
                handler_used="data_query_handler",
                execution_result=None,
                error_message=f"Data query error: {str(e)}"
            )
    
    def _route_data_update(self, llm_response: LLMResponse) -> RoutingResult:
        """Route data update operations."""
        # For now, return not implemented
        # This would be implemented when update handlers are available
        return RoutingResult(
            success=False,
            handler_used="data_update_handler",
            execution_result=None,
            error_message="Data update operations not yet implemented"
        )
    
    def _route_data_deletion(self, llm_response: LLMResponse) -> RoutingResult:
        """Route data deletion operations."""
        # For now, return not implemented
        # This would be implemented when deletion handlers are available
        return RoutingResult(
            success=False,
            handler_used="data_deletion_handler",
            execution_result=None,
            error_message="Data deletion operations not yet implemented"
        )
    
    def _route_chart_creation(self, llm_response: LLMResponse) -> RoutingResult:
        """Route chart creation operations."""
        try:
            if not self.excel_service.workbook:
                return RoutingResult(
                    success=False,
                    handler_used="visualization_operations",
                    execution_result=None,
                    error_message="No Excel file loaded"
                )
            
            parameters = llm_response.parameters
            
            result = self.visualization_operations.create_chart(
                workbook=self.excel_service.workbook,
                sheet_name=parameters.get('sheet_name', ''),
                data_range=parameters.get('data_range', ''),
                chart_type=parameters.get('chart_type'),
                title=parameters.get('title'),
                **{k: v for k, v in parameters.items() if k not in ['sheet_name', 'data_range', 'chart_type', 'title']}
            )
            
            return RoutingResult(
                success=result.get('success', False),
                handler_used="visualization_operations",
                execution_result=result
            )
            
        except Exception as e:
            return RoutingResult(
                success=False,
                handler_used="visualization_operations",
                execution_result=None,
                error_message=f"Chart creation error: {str(e)}"
            )
    
    def _route_chart_manipulation(self, llm_response: LLMResponse) -> RoutingResult:
        """Route chart manipulation operations."""
        try:
            if not self.chart_operations:
                return RoutingResult(
                    success=False,
                    handler_used="chart_operations",
                    execution_result=None,
                    error_message="Chart operations not available"
                )
            
            parameters = llm_response.parameters
            operation = llm_response.operation
            
            # Route to appropriate chart operation
            if operation == "shift_axis":
                result = self.chart_operations.shift_axis(
                    chart_id=parameters.get('chart_id', 'chart_1'),
                    axis=parameters.get('axis', 'x'),
                    shift_amount=parameters.get('shift_amount', 0.5)
                )
            elif operation == "transform_values":
                result = self.chart_operations.transform_values(
                    chart_id=parameters.get('chart_id', 'chart_1'),
                    axis=parameters.get('axis', 'y'),
                    operation=parameters.get('transform_function', 'subtract'),
                    value=parameters.get('transform_value', 1)
                )
            else:
                return RoutingResult(
                    success=False,
                    handler_used="chart_operations",
                    execution_result=None,
                    error_message=f"Chart operation '{operation}' not implemented"
                )
            
            return RoutingResult(
                success=result.get('success', False),
                handler_used="chart_operations",
                execution_result=result
            )
            
        except Exception as e:
            return RoutingResult(
                success=False,
                handler_used="chart_operations",
                execution_result=None,
                error_message=f"Chart manipulation error: {str(e)}"
            )
    
    def get_routing_statistics(self) -> Dict[str, Any]:
        """Get statistics about routing operations."""
        stats = {
            "routing_strategy": self.routing_strategy.value,
            "available_handlers": {
                "data_insertion": self.data_insertion_handler is not None,
                "data_query": self.data_query_handler is not None,
                "visualization": self.visualization_operations is not None,
                "chart_operations": self.chart_operations is not None
            },
            "routing_rules": {
                category: list(operations.keys()) 
                for category, operations in self._routing_rules.items()
            },
            "template_registry_stats": self.template_registry.get_registry_stats()
        }
        
        return stats
    
    def validate_routing_capability(self, intent: str, operation: str) -> Dict[str, Any]:
        """
        Validate if the router can handle a specific operation.
        
        Args:
            intent: Operation intent/category
            operation: Specific operation name
            
        Returns:
            Dict with validation results and available options
        """
        validation = {
            "can_route": False,
            "routing_method": None,
            "handler_available": False,
            "template_available": False,
            "error_message": None
        }
        
        # Check template registry
        if self.template_registry.is_operation_available(intent, operation):
            validation["template_available"] = True
            validation["can_route"] = True
            validation["routing_method"] = "template_registry"
        
        # Check direct handlers
        if intent in self._routing_rules and operation in self._routing_rules[intent]:
            validation["handler_available"] = True
            validation["can_route"] = True
            if not validation["routing_method"]:
                validation["routing_method"] = "direct_handler"
        
        if not validation["can_route"]:
            validation["error_message"] = f"No routing available for {intent}.{operation}"
        
        return validation