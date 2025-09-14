"""
Recovery manager for handling operation failures and system recovery.

This module provides comprehensive recovery mechanisms for various failure scenarios
including LLM connection failures, Excel file corruption, and operation rollbacks.
"""

import logging
import time
import shutil
from typing import Dict, Any, Optional, List, Callable
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from enum import Enum

from ..excel.excel_service import ExcelService, BackupInfo
from ..llm.ollama_service import OllamaService, OllamaConnectionError


class RecoveryAction(Enum):
    """Types of recovery actions available."""
    RETRY_OPERATION = "retry_operation"
    RESTORE_BACKUP = "restore_backup"
    RESTART_SERVICE = "restart_service"
    FALLBACK_MODE = "fallback_mode"
    USER_INTERVENTION = "user_intervention"
    ABORT_OPERATION = "abort_operation"


class RecoveryStrategy(Enum):
    """Recovery strategies for different failure types."""
    IMMEDIATE_RETRY = "immediate_retry"
    EXPONENTIAL_BACKOFF = "exponential_backoff"
    CIRCUIT_BREAKER = "circuit_breaker"
    GRACEFUL_DEGRADATION = "graceful_degradation"
    MANUAL_RECOVERY = "manual_recovery"


@dataclass
class RecoveryContext:
    """Context information for recovery operations."""
    operation_id: str
    failure_type: str
    error_message: str
    timestamp: datetime
    attempt_count: int
    max_attempts: int
    backup_path: Optional[str] = None
    original_file_path: Optional[str] = None
    recovery_data: Optional[Dict[str, Any]] = None


@dataclass
class RecoveryResult:
    """Result of a recovery operation."""
    success: bool
    action_taken: RecoveryAction
    message: str
    new_context: Optional[RecoveryContext] = None
    requires_user_action: bool = False
    user_instructions: Optional[List[str]] = None


class CircuitBreaker:
    """Circuit breaker pattern implementation for service failures."""
    
    def __init__(self, failure_threshold: int = 5, recovery_timeout: int = 60):
        """Initialize circuit breaker.
        
        Args:
            failure_threshold: Number of failures before opening circuit
            recovery_timeout: Seconds to wait before attempting recovery
        """
        self.failure_threshold = failure_threshold
        self.recovery_timeout = recovery_timeout
        self.failure_count = 0
        self.last_failure_time: Optional[datetime] = None
        self.state = "closed"  # closed, open, half-open
        self.logger = logging.getLogger(__name__)
    
    def call(self, func: Callable, *args, **kwargs) -> Any:
        """Execute function with circuit breaker protection.
        
        Args:
            func: Function to execute
            *args: Function arguments
            **kwargs: Function keyword arguments
            
        Returns:
            Function result
            
        Raises:
            Exception: If circuit is open or function fails
        """
        if self.state == "open":
            if self._should_attempt_reset():
                self.state = "half-open"
                self.logger.info("Circuit breaker attempting reset")
            else:
                raise Exception("Circuit breaker is open - service unavailable")
        
        try:
            result = func(*args, **kwargs)
            self._on_success()
            return result
        except Exception as e:
            self._on_failure()
            raise e
    
    def _should_attempt_reset(self) -> bool:
        """Check if enough time has passed to attempt reset."""
        if not self.last_failure_time:
            return True
        
        time_since_failure = datetime.now() - self.last_failure_time
        return time_since_failure.total_seconds() >= self.recovery_timeout
    
    def _on_success(self):
        """Handle successful operation."""
        self.failure_count = 0
        self.state = "closed"
        self.logger.info("Circuit breaker reset - service recovered")
    
    def _on_failure(self):
        """Handle failed operation."""
        self.failure_count += 1
        self.last_failure_time = datetime.now()
        
        if self.failure_count >= self.failure_threshold:
            self.state = "open"
            self.logger.warning(f"Circuit breaker opened after {self.failure_count} failures")


class RecoveryManager:
    """
    Comprehensive recovery manager for handling various failure scenarios.
    
    This class provides automated recovery mechanisms, backup restoration,
    and graceful degradation strategies for system failures.
    """
    
    def __init__(self, excel_service: ExcelService, ollama_service: OllamaService):
        """Initialize recovery manager.
        
        Args:
            excel_service: Excel service instance
            ollama_service: Ollama service instance
        """
        self.excel_service = excel_service
        self.ollama_service = ollama_service
        self.logger = logging.getLogger(__name__)
        
        # Circuit breakers for different services
        self.ollama_circuit_breaker = CircuitBreaker(failure_threshold=3, recovery_timeout=30)
        self.excel_circuit_breaker = CircuitBreaker(failure_threshold=5, recovery_timeout=10)
        
        # Recovery history
        self.recovery_history: List[RecoveryContext] = []
        self.max_history = 100
        
        # Recovery strategies mapping
        self.recovery_strategies = {
            "ollama_connection": RecoveryStrategy.EXPONENTIAL_BACKOFF,
            "excel_file_corruption": RecoveryStrategy.MANUAL_RECOVERY,
            "excel_permission": RecoveryStrategy.GRACEFUL_DEGRADATION,
            "operation_failure": RecoveryStrategy.IMMEDIATE_RETRY,
            "system_error": RecoveryStrategy.CIRCUIT_BREAKER
        }
    
    def handle_ollama_connection_failure(self, context: RecoveryContext) -> RecoveryResult:
        """Handle Ollama connection failures with retry logic.
        
        Args:
            context: Recovery context information
            
        Returns:
            RecoveryResult: Result of recovery attempt
        """
        self.logger.warning(f"Handling Ollama connection failure: {context.error_message}")
        
        # Check if we've exceeded max attempts
        if context.attempt_count >= context.max_attempts:
            return RecoveryResult(
                success=False,
                action_taken=RecoveryAction.ABORT_OPERATION,
                message="Maximum retry attempts exceeded for Ollama connection",
                requires_user_action=True,
                user_instructions=[
                    "Check if Ollama is running: 'ollama serve'",
                    "Verify Ollama installation: 'ollama --version'",
                    "Check network connectivity to Ollama endpoint",
                    "Restart Ollama service if necessary"
                ]
            )
        
        # Implement exponential backoff
        wait_time = min(2 ** context.attempt_count, 30)  # Cap at 30 seconds
        self.logger.info(f"Waiting {wait_time} seconds before retry attempt {context.attempt_count + 1}")
        time.sleep(wait_time)
        
        try:
            # Attempt to reconnect using circuit breaker
            success = self.ollama_circuit_breaker.call(self.ollama_service.initialize_connection)
            
            if success:
                return RecoveryResult(
                    success=True,
                    action_taken=RecoveryAction.RETRY_OPERATION,
                    message="Successfully reconnected to Ollama service"
                )
            else:
                # Update context for next attempt
                new_context = RecoveryContext(
                    operation_id=context.operation_id,
                    failure_type=context.failure_type,
                    error_message=context.error_message,
                    timestamp=datetime.now(),
                    attempt_count=context.attempt_count + 1,
                    max_attempts=context.max_attempts,
                    backup_path=context.backup_path,
                    original_file_path=context.original_file_path,
                    recovery_data=context.recovery_data
                )
                
                return RecoveryResult(
                    success=False,
                    action_taken=RecoveryAction.RETRY_OPERATION,
                    message=f"Retry attempt {context.attempt_count + 1} failed",
                    new_context=new_context
                )
        
        except Exception as e:
            self.logger.error(f"Recovery attempt failed: {str(e)}")
            
            # Update context for next attempt
            new_context = RecoveryContext(
                operation_id=context.operation_id,
                failure_type=context.failure_type,
                error_message=str(e),
                timestamp=datetime.now(),
                attempt_count=context.attempt_count + 1,
                max_attempts=context.max_attempts,
                backup_path=context.backup_path,
                original_file_path=context.original_file_path,
                recovery_data=context.recovery_data
            )
            
            return RecoveryResult(
                success=False,
                action_taken=RecoveryAction.RETRY_OPERATION,
                message=f"Connection attempt failed: {str(e)}",
                new_context=new_context
            )
    
    def handle_excel_file_error(self, context: RecoveryContext) -> RecoveryResult:
        """Handle Excel file errors with appropriate recovery strategies.
        
        Args:
            context: Recovery context information
            
        Returns:
            RecoveryResult: Result of recovery attempt
        """
        self.logger.warning(f"Handling Excel file error: {context.error_message}")
        
        error_message = context.error_message.lower()
        
        # Handle permission errors
        if "permission" in error_message or "access" in error_message:
            return self._handle_permission_error(context)
        
        # Handle file corruption
        elif "corrupt" in error_message or "invalid" in error_message or "format" in error_message:
            return self._handle_file_corruption(context)
        
        # Handle file not found
        elif "not found" in error_message or "no such file" in error_message:
            return self._handle_file_not_found(context)
        
        # Generic file error handling
        else:
            return self._handle_generic_file_error(context)
    
    def _handle_permission_error(self, context: RecoveryContext) -> RecoveryResult:
        """Handle file permission errors."""
        self.logger.info("Attempting to resolve permission error")
        
        # Try to detect if file is open in Excel
        file_path = context.original_file_path or context.recovery_data.get("file_path")
        if file_path:
            temp_file = f"{file_path}.tmp"
            try:
                # Try to create a temporary file to test write permissions
                shutil.copy2(file_path, temp_file)
                Path(temp_file).unlink()  # Clean up
                
                # If we can copy, the issue might be Excel having the file open
                return RecoveryResult(
                    success=False,
                    action_taken=RecoveryAction.USER_INTERVENTION,
                    message="File appears to be open in Excel or another application",
                    requires_user_action=True,
                    user_instructions=[
                        "Close the Excel file if it's currently open",
                        "Check if any other applications are using the file",
                        "Ensure you have write permissions to the file",
                        "Try running the application as administrator if needed"
                    ]
                )
            
            except Exception:
                return RecoveryResult(
                    success=False,
                    action_taken=RecoveryAction.USER_INTERVENTION,
                    message="Insufficient permissions to access the file",
                    requires_user_action=True,
                    user_instructions=[
                        "Check file permissions and ensure you have write access",
                        "Try running the application as administrator",
                        "Move the file to a location with appropriate permissions",
                        "Contact your system administrator if needed"
                    ]
                )
        
        return RecoveryResult(
            success=False,
            action_taken=RecoveryAction.USER_INTERVENTION,
            message="Permission error - user intervention required",
            requires_user_action=True,
            user_instructions=["Check file permissions and access rights"]
        )
    
    def _handle_file_corruption(self, context: RecoveryContext) -> RecoveryResult:
        """Handle file corruption with backup restoration."""
        self.logger.info("Attempting to recover from file corruption")
        
        # Try to restore from backup
        if context.backup_path and Path(context.backup_path).exists():
            try:
                success = self.excel_service.restore_from_backup(context.backup_path)
                if success:
                    return RecoveryResult(
                        success=True,
                        action_taken=RecoveryAction.RESTORE_BACKUP,
                        message=f"Successfully restored from backup: {context.backup_path}"
                    )
            except Exception as e:
                self.logger.error(f"Failed to restore from backup: {str(e)}")
        
        # Try to find and restore from most recent backup
        if context.original_file_path:
            original_name = Path(context.original_file_path).stem
            backups = self.excel_service.get_backup_list(original_name)
            
            for backup in backups[:3]:  # Try up to 3 most recent backups
                try:
                    success = self.excel_service.restore_from_backup(backup.file_path)
                    if success:
                        return RecoveryResult(
                            success=True,
                            action_taken=RecoveryAction.RESTORE_BACKUP,
                            message=f"Successfully restored from backup: {backup.file_path}"
                        )
                except Exception as e:
                    self.logger.warning(f"Failed to restore from backup {backup.file_path}: {str(e)}")
                    continue
        
        # If backup restoration fails, suggest manual recovery
        return RecoveryResult(
            success=False,
            action_taken=RecoveryAction.USER_INTERVENTION,
            message="File appears to be corrupted and backup restoration failed",
            requires_user_action=True,
            user_instructions=[
                "Try opening the file in Excel to see if it can be repaired",
                "Use Excel's built-in repair functionality",
                "Restore from a known good backup manually",
                "Contact support if the file contains critical data"
            ]
        )
    
    def _handle_file_not_found(self, context: RecoveryContext) -> RecoveryResult:
        """Handle file not found errors."""
        file_path = context.original_file_path or context.recovery_data.get("file_path", "")
        
        # Check if file exists in common alternative locations
        if file_path:
            path_obj = Path(file_path)
            alternative_paths = [
                path_obj.parent / path_obj.name,  # Same directory
                Path.cwd() / path_obj.name,       # Current working directory
                Path.home() / "Documents" / path_obj.name,  # Documents folder
                Path.home() / "Desktop" / path_obj.name     # Desktop
            ]
            
            for alt_path in alternative_paths:
                if alt_path.exists() and alt_path != Path(file_path):
                    return RecoveryResult(
                        success=False,
                        action_taken=RecoveryAction.USER_INTERVENTION,
                        message=f"File not found at specified location, but found at: {alt_path}",
                        requires_user_action=True,
                        user_instructions=[
                            f"File found at alternative location: {alt_path}",
                            "Update the file path in your command",
                            "Or move the file to the expected location"
                        ]
                    )
        
        return RecoveryResult(
            success=False,
            action_taken=RecoveryAction.USER_INTERVENTION,
            message="File not found",
            requires_user_action=True,
            user_instructions=[
                "Verify the file path is correct",
                "Check if the file has been moved or renamed",
                "Ensure the file exists and is accessible",
                "Use the full absolute path to the file"
            ]
        )
    
    def _handle_generic_file_error(self, context: RecoveryContext) -> RecoveryResult:
        """Handle generic file errors."""
        return RecoveryResult(
            success=False,
            action_taken=RecoveryAction.USER_INTERVENTION,
            message=f"Excel file error: {context.error_message}",
            requires_user_action=True,
            user_instructions=[
                "Check if the file is in a supported format (.xlsx, .xls, .csv)",
                "Verify the file is not corrupted",
                "Ensure you have appropriate permissions",
                "Try opening the file in Excel to verify it's valid"
            ]
        )
    
    def handle_operation_failure(self, context: RecoveryContext) -> RecoveryResult:
        """Handle general operation failures with rollback capability.
        
        Args:
            context: Recovery context information
            
        Returns:
            RecoveryResult: Result of recovery attempt
        """
        self.logger.warning(f"Handling operation failure: {context.error_message}")
        
        # If we have a backup, offer to restore it
        if context.backup_path and Path(context.backup_path).exists():
            try:
                success = self.excel_service.restore_from_backup(context.backup_path)
                if success:
                    return RecoveryResult(
                        success=True,
                        action_taken=RecoveryAction.RESTORE_BACKUP,
                        message="Operation failed - successfully restored from backup"
                    )
            except Exception as e:
                self.logger.error(f"Failed to restore backup after operation failure: {str(e)}")
        
        # If immediate retry is appropriate and we haven't exceeded attempts
        if context.attempt_count < context.max_attempts:
            new_context = RecoveryContext(
                operation_id=context.operation_id,
                failure_type=context.failure_type,
                error_message=context.error_message,
                timestamp=datetime.now(),
                attempt_count=context.attempt_count + 1,
                max_attempts=context.max_attempts,
                backup_path=context.backup_path,
                original_file_path=context.original_file_path,
                recovery_data=context.recovery_data
            )
            
            return RecoveryResult(
                success=False,
                action_taken=RecoveryAction.RETRY_OPERATION,
                message=f"Operation failed - will retry (attempt {context.attempt_count + 1})",
                new_context=new_context
            )
        
        # Max attempts exceeded
        return RecoveryResult(
            success=False,
            action_taken=RecoveryAction.ABORT_OPERATION,
            message="Operation failed after maximum retry attempts",
            requires_user_action=True,
            user_instructions=[
                "Review the error details and try a different approach",
                "Check if the operation parameters are correct",
                "Ensure the Excel file is in a valid state",
                "Contact support if the problem persists"
            ]
        )
    
    def create_recovery_context(self, operation_id: str, failure_type: str, 
                             error_message: str, max_attempts: int = 3,
                             backup_path: Optional[str] = None,
                             original_file_path: Optional[str] = None,
                             recovery_data: Optional[Dict[str, Any]] = None) -> RecoveryContext:
        """Create a recovery context for tracking recovery operations.
        
        Args:
            operation_id: Unique identifier for the operation
            failure_type: Type of failure that occurred
            error_message: Error message from the failure
            max_attempts: Maximum number of recovery attempts
            backup_path: Path to backup file if available
            original_file_path: Path to original file
            recovery_data: Additional recovery data
            
        Returns:
            RecoveryContext: Created recovery context
        """
        context = RecoveryContext(
            operation_id=operation_id,
            failure_type=failure_type,
            error_message=error_message,
            timestamp=datetime.now(),
            attempt_count=0,
            max_attempts=max_attempts,
            backup_path=backup_path,
            original_file_path=original_file_path,
            recovery_data=recovery_data or {}
        )
        
        # Add to history
        self.recovery_history.append(context)
        if len(self.recovery_history) > self.max_history:
            self.recovery_history = self.recovery_history[-self.max_history:]
        
        return context
    
    def get_recovery_statistics(self) -> Dict[str, Any]:
        """Get statistics about recovery operations.
        
        Returns:
            Dict containing recovery statistics
        """
        if not self.recovery_history:
            return {"total_recoveries": 0}
        
        # Count by failure type
        failure_counts = {}
        recent_failures = []
        
        for context in self.recovery_history:
            failure_type = context.failure_type
            failure_counts[failure_type] = failure_counts.get(failure_type, 0) + 1
            
            # Recent failures (last 24 hours)
            if datetime.now() - context.timestamp < timedelta(hours=24):
                recent_failures.append({
                    "operation_id": context.operation_id,
                    "failure_type": context.failure_type,
                    "error_message": context.error_message,
                    "timestamp": context.timestamp.isoformat(),
                    "attempt_count": context.attempt_count
                })
        
        return {
            "total_recoveries": len(self.recovery_history),
            "by_failure_type": failure_counts,
            "recent_failures_24h": recent_failures,
            "circuit_breaker_states": {
                "ollama": self.ollama_circuit_breaker.state,
                "excel": self.excel_circuit_breaker.state
            }
        }
    
    def reset_circuit_breakers(self):
        """Reset all circuit breakers to closed state."""
        self.ollama_circuit_breaker.failure_count = 0
        self.ollama_circuit_breaker.state = "closed"
        self.ollama_circuit_breaker.last_failure_time = None
        
        self.excel_circuit_breaker.failure_count = 0
        self.excel_circuit_breaker.state = "closed"
        self.excel_circuit_breaker.last_failure_time = None
        
        self.logger.info("All circuit breakers reset")