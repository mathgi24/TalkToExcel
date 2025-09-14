"""
Comprehensive logging and audit trail system for Excel-LLM operations.

This module provides detailed logging, audit trails, and operation tracking
for compliance, debugging, and system monitoring purposes.
"""

import logging
import json
import os
from typing import Dict, Any, Optional, List
from dataclasses import dataclass, asdict
from datetime import datetime, timedelta
from pathlib import Path
from enum import Enum
import threading
from logging.handlers import RotatingFileHandler

# from ..config.config_manager import config


class AuditEventType(Enum):
    """Types of audit events."""
    OPERATION_START = "operation_start"
    OPERATION_SUCCESS = "operation_success"
    OPERATION_FAILURE = "operation_failure"
    FILE_ACCESS = "file_access"
    BACKUP_CREATED = "backup_created"
    BACKUP_RESTORED = "backup_restored"
    LLM_REQUEST = "llm_request"
    LLM_RESPONSE = "llm_response"
    SAFETY_VIOLATION = "safety_violation"
    RECOVERY_ATTEMPT = "recovery_attempt"
    USER_CONFIRMATION = "user_confirmation"
    SYSTEM_ERROR = "system_error"
    CONFIG_CHANGE = "config_change"


class LogLevel(Enum):
    """Log levels for audit events."""
    DEBUG = "DEBUG"
    INFO = "INFO"
    WARNING = "WARNING"
    ERROR = "ERROR"
    CRITICAL = "CRITICAL"


@dataclass
class AuditEvent:
    """Audit event data structure."""
    event_id: str
    event_type: AuditEventType
    timestamp: datetime
    user_id: Optional[str]
    session_id: Optional[str]
    operation_id: Optional[str]
    component: str
    action: str
    details: Dict[str, Any]
    result: Optional[str] = None
    error_message: Optional[str] = None
    file_path: Optional[str] = None
    backup_path: Optional[str] = None
    duration_ms: Optional[int] = None
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert audit event to dictionary."""
        data = asdict(self)
        data['timestamp'] = self.timestamp.isoformat()
        data['event_type'] = self.event_type.value
        return data
    
    def to_json(self) -> str:
        """Convert audit event to JSON string."""
        return json.dumps(self.to_dict(), indent=2)


class AuditLogger:
    """
    Comprehensive audit logging system for tracking all system operations.
    
    This class provides detailed audit trails, structured logging, and
    compliance-ready operation tracking.
    """
    
    def __init__(self):
        """Initialize audit logger."""
        # Fallback config for now
        self.config = {
            'file': './logs/excel_llm.log',
            'level': 'INFO',
            'max_file_size': '10MB',
            'backup_count': 5,
            'console_output': True
        }
        
        self.audit_events: List[AuditEvent] = []
        self.max_events_in_memory = 1000
        self._lock = threading.Lock()
        
        # Setup loggers
        self._setup_loggers()
        
        # Generate session ID
        self.session_id = self._generate_session_id()
        
        # Event counter for unique IDs
        self._event_counter = 0
    
    def _setup_loggers(self):
        """Setup structured logging with file rotation."""
        log_dir = Path(self.config.get('file', './logs/excel_llm.log')).parent
        log_dir.mkdir(parents=True, exist_ok=True)
        
        # Main application logger
        self.app_logger = logging.getLogger('excel_llm_app')
        self.app_logger.setLevel(getattr(logging, self.config.get('level', 'INFO')))
        
        # Audit trail logger (separate file)
        self.audit_logger = logging.getLogger('excel_llm_audit')
        self.audit_logger.setLevel(logging.INFO)
        
        # Performance logger
        self.perf_logger = logging.getLogger('excel_llm_performance')
        self.perf_logger.setLevel(logging.INFO)
        
        # Setup file handlers with rotation
        max_bytes = self._parse_size(self.config.get('max_file_size', '10MB'))
        backup_count = self.config.get('backup_count', 5)
        
        # Application log handler
        app_handler = RotatingFileHandler(
            log_dir / 'application.log',
            maxBytes=max_bytes,
            backupCount=backup_count
        )
        app_formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        app_handler.setFormatter(app_formatter)
        self.app_logger.addHandler(app_handler)
        
        # Audit log handler (JSON format)
        audit_handler = RotatingFileHandler(
            log_dir / 'audit.log',
            maxBytes=max_bytes,
            backupCount=backup_count
        )
        audit_formatter = logging.Formatter('%(message)s')
        audit_handler.setFormatter(audit_formatter)
        self.audit_logger.addHandler(audit_handler)
        
        # Performance log handler
        perf_handler = RotatingFileHandler(
            log_dir / 'performance.log',
            maxBytes=max_bytes,
            backupCount=backup_count
        )
        perf_formatter = logging.Formatter(
            '%(asctime)s - PERF - %(message)s'
        )
        perf_handler.setFormatter(perf_formatter)
        self.perf_logger.addHandler(perf_handler)
        
        # Console handler for development
        if self.config.get('console_output', True):
            console_handler = logging.StreamHandler()
            console_formatter = logging.Formatter(
                '%(asctime)s - %(levelname)s - %(message)s'
            )
            console_handler.setFormatter(console_formatter)
            self.app_logger.addHandler(console_handler)
    
    def _parse_size(self, size_str: str) -> int:
        """Parse size string (e.g., '10MB') to bytes."""
        size_str = size_str.upper()
        if size_str.endswith('KB'):
            return int(size_str[:-2]) * 1024
        elif size_str.endswith('MB'):
            return int(size_str[:-2]) * 1024 * 1024
        elif size_str.endswith('GB'):
            return int(size_str[:-2]) * 1024 * 1024 * 1024
        else:
            return int(size_str)
    
    def _generate_session_id(self) -> str:
        """Generate unique session ID."""
        return f"session_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{os.getpid()}"
    
    def _generate_event_id(self) -> str:
        """Generate unique event ID."""
        with self._lock:
            self._event_counter += 1
            return f"evt_{self.session_id}_{self._event_counter:06d}"
    
    def log_audit_event(self, event_type: AuditEventType, component: str, 
                       action: str, details: Dict[str, Any],
                       user_id: Optional[str] = None,
                       operation_id: Optional[str] = None,
                       result: Optional[str] = None,
                       error_message: Optional[str] = None,
                       file_path: Optional[str] = None,
                       backup_path: Optional[str] = None,
                       duration_ms: Optional[int] = None) -> str:
        """Log an audit event."""
        event_id = self._generate_event_id()
        
        # Create audit event
        event = AuditEvent(
            event_id=event_id,
            event_type=event_type,
            timestamp=datetime.now(),
            user_id=user_id,
            session_id=self.session_id,
            operation_id=operation_id,
            component=component,
            action=action,
            details=details.copy(),  # Make a copy to avoid mutations
            result=result,
            error_message=error_message,
            file_path=file_path,
            backup_path=backup_path,
            duration_ms=duration_ms
        )
        
        # Add to in-memory storage
        with self._lock:
            self.audit_events.append(event)
            if len(self.audit_events) > self.max_events_in_memory:
                self.audit_events = self.audit_events[-self.max_events_in_memory:]
        
        # Log to audit file
        self.audit_logger.info(event.to_json())
        
        return event_id
    
    def get_audit_statistics(self) -> Dict[str, Any]:
        """Get audit statistics."""
        with self._lock:
            events = self.audit_events.copy()
        
        if not events:
            return {"total_events": 0}
        
        # Count by event type
        event_type_counts = {}
        component_counts = {}
        
        for event in events:
            event_type = event.event_type.value
            component = event.component
            
            event_type_counts[event_type] = event_type_counts.get(event_type, 0) + 1
            component_counts[component] = component_counts.get(component, 0) + 1
        
        return {
            "total_events": len(events),
            "session_id": self.session_id,
            "by_event_type": event_type_counts,
            "by_component": component_counts
        }