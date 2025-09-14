"""
Error handling and user feedback systems for command processing.

This module provides comprehensive error handling, user feedback generation,
and recovery mechanisms for the command processing pipeline.
"""

import logging
import traceback
from typing import Dict, Any, Optional, List, Tuple
from dataclasses import dataclass
from enum import Enum
from datetime import datetime

from llm.ollama_service import OllamaConnectionError

# Add src directory to Python path for imports
import sys
from pathlib import Path
current_dir = Path(__file__).parent
src_dir = current_dir.parent if current_dir.name != 'src' else current_dir
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))



class ErrorCategory(Enum):
    """Categories of errors that can occur."""
    LLM_CONNECTION = "llm_connection"
    COMMAND_PARSING = "command_parsing"
    OPERATION_EXECUTION = "operation_execution"
    SAFETY_VIOLATION = "safety_violation"
    FILE_ACCESS = "file_access"
    VALIDATION = "validation"
    SYSTEM = "system"
    USER_INPUT = "user_input"


class ErrorSeverity(Enum):
    """Severity levels for errors."""
    LOW = "low"
    MEDIUM = "medium"
    HIGH = "high"
    CRITICAL = "critical"


@dataclass
class ErrorInfo:
    """Comprehensive error information."""
    category: ErrorCategory
    severity: ErrorSeverity
    message: str
    technical_details: Optional[str] = None
    user_message: Optional[str] = None
    suggestions: Optional[List[str]] = None
    recovery_actions: Optional[List[str]] = None
    timestamp: Optional[datetime] = None
    
    def __post_init__(self):
        if self.timestamp is None:
            self.timestamp = datetime.now()
        if self.suggestions is None:
            self.suggestions = []
        if self.recovery_actions is None:
            self.recovery_actions = []


@dataclass
class FeedbackMessage:
    """User feedback message."""
    message: str
    message_type: str  # "success", "warning", "error", "info"
    details: Optional[str] = None
    actions: Optional[List[str]] = None
    
    def __post_init__(self):
        if self.actions is None:
            self.actions = []


class ErrorHandler:
    """
    Comprehensive error handling and user feedback system.
    
    This class provides centralized error handling, user-friendly error messages,
    and recovery suggestions for all types of errors in the system.
    """
    
    def __init__(self):
        """Initialize error handler."""
        self.logger = logging.getLogger(__name__)
        self._error_history: List[ErrorInfo] = []
        self._max_history = 100
        
        # Define error patterns and their handling
        self._error_patterns = self._build_error_patterns()
    
    def _build_error_patterns(self) -> Dict[str, Dict[str, Any]]:
        """Build patterns for recognizing and handling different error types."""
        return {
            # LLM Connection Errors
            "connection_refused": {
                "category": ErrorCategory.LLM_CONNECTION,
                "severity": ErrorSeverity.HIGH,
                "keywords": ["connection refused", "connection failed", "timeout"],
                "user_message": "Cannot connect to the AI service. Please check if Ollama is running.",
                "suggestions": [
                    "Start Ollama service: 'ollama serve'",
                    "Check if Ollama is installed correctly",
                    "Verify the endpoint configuration"
                ],
                "recovery_actions": ["retry_connection", "check_service_status"]
            },
            
            "model_not_found": {
                "category": ErrorCategory.LLM_CONNECTION,
                "severity": ErrorSeverity.HIGH,
                "keywords": ["model not found", "model unavailable"],
                "user_message": "The AI model is not available. Please download the required model.",
                "suggestions": [
                    "Download the model: 'ollama pull mistral:7b-instruct'",
                    "Check available models: 'ollama list'",
                    "Verify model name in configuration"
                ],
                "recovery_actions": ["download_model", "list_models"]
            },
            
            # File Access Errors
            "file_not_found": {
                "category": ErrorCategory.FILE_ACCESS,
                "severity": ErrorSeverity.MEDIUM,
                "keywords": ["file not found", "no such file"],
                "user_message": "The Excel file could not be found.",
                "suggestions": [
                    "Check if the file path is correct",
                    "Ensure the file exists and is accessible",
                    "Try using an absolute file path"
                ],
                "recovery_actions": ["verify_file_path", "browse_for_file"]
            },
            
            "permission_denied": {
                "category": ErrorCategory.FILE_ACCESS,
                "severity": ErrorSeverity.MEDIUM,
                "keywords": ["permission denied", "access denied"],
                "user_message": "Cannot access the Excel file due to permission restrictions.",
                "suggestions": [
                    "Check file permissions",
                    "Close the file if it's open in Excel",
                    "Run with administrator privileges if needed"
                ],
                "recovery_actions": ["check_permissions", "close_excel"]
            },
            
            # Operation Execution Errors
            "operation_not_implemented": {
                "category": ErrorCategory.OPERATION_EXECUTION,
                "severity": ErrorSeverity.LOW,
                "keywords": ["not implemented", "not available"],
                "user_message": "This operation is not yet implemented.",
                "suggestions": [
                    "Try a similar operation that is available",
                    "Check the list of supported operations",
                    "Consider using a different approach"
                ],
                "recovery_actions": ["list_available_operations", "suggest_alternatives"]
            },
            
            # Safety Violations
            "operation_blocked": {
                "category": ErrorCategory.SAFETY_VIOLATION,
                "severity": ErrorSeverity.MEDIUM,
                "keywords": ["blocked", "not allowed", "safety"],
                "user_message": "This operation was blocked for safety reasons.",
                "suggestions": [
                    "Try a more specific operation",
                    "Break down the operation into smaller parts",
                    "Review the safety guidelines"
                ],
                "recovery_actions": ["suggest_safe_alternatives", "explain_safety_rules"]
            },
            
            # Validation Errors
            "invalid_parameters": {
                "category": ErrorCategory.VALIDATION,
                "severity": ErrorSeverity.LOW,
                "keywords": ["invalid", "validation failed", "parameter"],
                "user_message": "The operation parameters are invalid.",
                "suggestions": [
                    "Check the parameter values",
                    "Ensure all required parameters are provided",
                    "Verify data types and formats"
                ],
                "recovery_actions": ["validate_parameters", "show_parameter_help"]
            }
        }
    
    def handle_error(self, error: Exception, context: Optional[Dict[str, Any]] = None) -> ErrorInfo:
        """
        Handle an error and generate comprehensive error information.
        
        Args:
            error: The exception that occurred
            context: Additional context about where the error occurred
            
        Returns:
            ErrorInfo: Comprehensive error information
        """
        error_str = str(error).lower()
        error_type = type(error).__name__
        
        # Classify the error
        error_info = self._classify_error(error, error_str, context)
        
        # Add to history
        self._add_to_history(error_info)
        
        # Log the error
        self._log_error(error_info, error)
        
        return error_info
    
    def _classify_error(self, error: Exception, error_str: str, context: Optional[Dict[str, Any]]) -> ErrorInfo:
        """Classify error and generate appropriate error information."""
        # Check for specific error types first
        if isinstance(error, OllamaConnectionError):
            return self._handle_ollama_error(error, error_str)
        elif isinstance(error, FileNotFoundError):
            return self._handle_file_error(error, error_str)
        elif isinstance(error, PermissionError):
            return self._handle_permission_error(error, error_str)
        elif isinstance(error, NotImplementedError):
            return self._handle_not_implemented_error(error, error_str)
        
        # Pattern-based classification
        for pattern_name, pattern_info in self._error_patterns.items():
            if any(keyword in error_str for keyword in pattern_info["keywords"]):
                return ErrorInfo(
                    category=pattern_info["category"],
                    severity=pattern_info["severity"],
                    message=str(error),
                    user_message=pattern_info["user_message"],
                    suggestions=pattern_info["suggestions"].copy(),
                    recovery_actions=pattern_info["recovery_actions"].copy(),
                    technical_details=self._get_technical_details(error, context)
                )
        
        # Default error handling
        return ErrorInfo(
            category=ErrorCategory.SYSTEM,
            severity=ErrorSeverity.MEDIUM,
            message=str(error),
            user_message="An unexpected error occurred while processing your request.",
            suggestions=[
                "Try rephrasing your command",
                "Check if all required information is provided",
                "Contact support if the problem persists"
            ],
            technical_details=self._get_technical_details(error, context)
        )
    
    def _handle_ollama_error(self, error: OllamaConnectionError, error_str: str) -> ErrorInfo:
        """Handle Ollama-specific errors."""
        if "connection refused" in error_str or "connection failed" in error_str:
            return ErrorInfo(
                category=ErrorCategory.LLM_CONNECTION,
                severity=ErrorSeverity.HIGH,
                message=str(error),
                user_message="Cannot connect to Ollama. Please ensure Ollama is running.",
                suggestions=[
                    "Start Ollama: 'ollama serve'",
                    "Check if Ollama is installed: 'ollama --version'",
                    "Verify the endpoint in configuration"
                ],
                recovery_actions=["start_ollama", "check_installation"]
            )
        elif "model" in error_str and "not found" in error_str:
            return ErrorInfo(
                category=ErrorCategory.LLM_CONNECTION,
                severity=ErrorSeverity.HIGH,
                message=str(error),
                user_message="The required AI model is not available.",
                suggestions=[
                    "Download the model: 'ollama pull mistral:7b-instruct'",
                    "List available models: 'ollama list'",
                    "Check model name in configuration"
                ],
                recovery_actions=["download_model", "list_models"]
            )
        else:
            return ErrorInfo(
                category=ErrorCategory.LLM_CONNECTION,
                severity=ErrorSeverity.MEDIUM,
                message=str(error),
                user_message="AI service error occurred.",
                suggestions=["Try again in a moment", "Check Ollama service status"],
                recovery_actions=["retry_connection"]
            )
    
    def _handle_file_error(self, error: FileNotFoundError, error_str: str) -> ErrorInfo:
        """Handle file not found errors."""
        return ErrorInfo(
            category=ErrorCategory.FILE_ACCESS,
            severity=ErrorSeverity.MEDIUM,
            message=str(error),
            user_message="The Excel file could not be found.",
            suggestions=[
                "Check if the file path is correct",
                "Ensure the file exists",
                "Try using the full file path"
            ],
            recovery_actions=["verify_path", "browse_file"]
        )
    
    def _handle_permission_error(self, error: PermissionError, error_str: str) -> ErrorInfo:
        """Handle permission errors."""
        return ErrorInfo(
            category=ErrorCategory.FILE_ACCESS,
            severity=ErrorSeverity.MEDIUM,
            message=str(error),
            user_message="Cannot access the file due to permission restrictions.",
            suggestions=[
                "Close the file if it's open in Excel",
                "Check file permissions",
                "Try running as administrator"
            ],
            recovery_actions=["close_file", "check_permissions"]
        )
    
    def _handle_not_implemented_error(self, error: NotImplementedError, error_str: str) -> ErrorInfo:
        """Handle not implemented errors."""
        return ErrorInfo(
            category=ErrorCategory.OPERATION_EXECUTION,
            severity=ErrorSeverity.LOW,
            message=str(error),
            user_message="This operation is not yet available.",
            suggestions=[
                "Try a similar operation",
                "Check available operations",
                "Use a different approach"
            ],
            recovery_actions=["list_operations", "suggest_alternatives"]
        )
    
    def _get_technical_details(self, error: Exception, context: Optional[Dict[str, Any]]) -> str:
        """Get technical details for debugging."""
        details = [
            f"Error Type: {type(error).__name__}",
            f"Error Message: {str(error)}"
        ]
        
        if context:
            details.append(f"Context: {context}")
        
        # Add stack trace for debugging
        details.append(f"Stack Trace: {traceback.format_exc()}")
        
        return "\n".join(details)
    
    def _add_to_history(self, error_info: ErrorInfo):
        """Add error to history with size limit."""
        self._error_history.append(error_info)
        
        # Maintain history size limit
        if len(self._error_history) > self._max_history:
            self._error_history = self._error_history[-self._max_history:]
    
    def _log_error(self, error_info: ErrorInfo, original_error: Exception):
        """Log error with appropriate level."""
        log_message = f"[{error_info.category.value}] {error_info.message}"
        
        if error_info.severity == ErrorSeverity.CRITICAL:
            self.logger.critical(log_message, exc_info=original_error)
        elif error_info.severity == ErrorSeverity.HIGH:
            self.logger.error(log_message, exc_info=original_error)
        elif error_info.severity == ErrorSeverity.MEDIUM:
            self.logger.warning(log_message)
        else:
            self.logger.info(log_message)
    
    def generate_user_feedback(self, error_info: ErrorInfo) -> FeedbackMessage:
        """Generate user-friendly feedback message."""
        message_type = self._get_message_type(error_info.severity)
        
        # Use user message if available, otherwise use technical message
        main_message = error_info.user_message or error_info.message
        
        # Prepare actions
        actions = []
        if error_info.suggestions:
            actions.extend(error_info.suggestions)
        
        return FeedbackMessage(
            message=main_message,
            message_type=message_type,
            details=error_info.technical_details if error_info.severity in [ErrorSeverity.HIGH, ErrorSeverity.CRITICAL] else None,
            actions=actions
        )
    
    def _get_message_type(self, severity: ErrorSeverity) -> str:
        """Convert severity to message type."""
        severity_mapping = {
            ErrorSeverity.LOW: "warning",
            ErrorSeverity.MEDIUM: "error",
            ErrorSeverity.HIGH: "error",
            ErrorSeverity.CRITICAL: "error"
        }
        return severity_mapping.get(severity, "error")
    
    def get_recovery_suggestions(self, error_category: ErrorCategory) -> List[str]:
        """Get recovery suggestions for a specific error category."""
        suggestions = []
        
        for pattern_info in self._error_patterns.values():
            if pattern_info["category"] == error_category:
                suggestions.extend(pattern_info["suggestions"])
        
        # Remove duplicates while preserving order
        return list(dict.fromkeys(suggestions))
    
    def get_error_statistics(self) -> Dict[str, Any]:
        """Get statistics about errors that have occurred."""
        if not self._error_history:
            return {"total_errors": 0}
        
        # Count by category
        category_counts = {}
        severity_counts = {}
        
        for error in self._error_history:
            category = error.category.value
            severity = error.severity.value
            
            category_counts[category] = category_counts.get(category, 0) + 1
            severity_counts[severity] = severity_counts.get(severity, 0) + 1
        
        # Recent errors (last 10)
        recent_errors = [
            {
                "category": error.category.value,
                "severity": error.severity.value,
                "message": error.message,
                "timestamp": error.timestamp.isoformat() if error.timestamp else None
            }
            for error in self._error_history[-10:]
        ]
        
        return {
            "total_errors": len(self._error_history),
            "by_category": category_counts,
            "by_severity": severity_counts,
            "recent_errors": recent_errors
        }
    
    def clear_error_history(self):
        """Clear the error history."""
        self._error_history.clear()
        self.logger.info("Error history cleared")


class UserFeedbackGenerator:
    """Generates user-friendly feedback messages for various scenarios."""
    
    def __init__(self):
        """Initialize feedback generator."""
        self.logger = logging.getLogger(__name__)
    
    def generate_success_message(self, operation: str, details: Optional[Dict[str, Any]] = None) -> FeedbackMessage:
        """Generate success feedback message."""
        base_messages = {
            "data_creation": "Data added successfully!",
            "data_query": "Data retrieved successfully!",
            "chart_creation": "Chart created successfully!",
            "chart_manipulation": "Chart updated successfully!"
        }
        
        message = base_messages.get(operation, "Operation completed successfully!")
        
        # Add details if available
        if details:
            detail_parts = []
            if "affected_rows" in details:
                detail_parts.append(f"{details['affected_rows']} rows affected")
            if "chart_id" in details:
                detail_parts.append(f"Chart ID: {details['chart_id']}")
            if "row_count" in details:
                detail_parts.append(f"{details['row_count']} records found")
            
            if detail_parts:
                message += f" ({', '.join(detail_parts)})"
        
        return FeedbackMessage(
            message=message,
            message_type="success",
            details=str(details) if details else None
        )
    
    def generate_warning_message(self, warning: str, suggestions: Optional[List[str]] = None) -> FeedbackMessage:
        """Generate warning feedback message."""
        return FeedbackMessage(
            message=f"Warning: {warning}",
            message_type="warning",
            actions=suggestions or []
        )
    
    def generate_info_message(self, info: str, actions: Optional[List[str]] = None) -> FeedbackMessage:
        """Generate informational feedback message."""
        return FeedbackMessage(
            message=info,
            message_type="info",
            actions=actions or []
        )
    
    def generate_confirmation_message(self, operation: str, details: Dict[str, Any]) -> FeedbackMessage:
        """Generate confirmation request message."""
        risk_level = details.get("risk_level", "medium")
        operation_desc = details.get("operation", operation)
        
        message = f"Confirm {risk_level} risk operation: {operation_desc}"
        
        actions = [
            "Type 'yes' to proceed",
            "Type 'no' to cancel",
            "Review the operation details below"
        ]
        
        return FeedbackMessage(
            message=message,
            message_type="warning",
            details=str(details),
            actions=actions
        )