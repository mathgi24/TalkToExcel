"""
Risk assessment classifier for Excel operations.

Evaluates the risk level of operations based on their potential impact
on data integrity and spreadsheet structure.
"""

from enum import Enum
from typing import Dict, List, Any, Optional
from dataclasses import dataclass


class RiskLevel(Enum):
    """Risk levels for operations."""
    LOW = "low"
    MEDIUM = "medium"
    HIGH = "high"
    DANGEROUS = "dangerous"


@dataclass
class RiskAssessment:
    """Result of risk assessment for an operation."""
    level: RiskLevel
    score: float  # 0.0 to 1.0
    reasons: List[str]
    blocked: bool = False
    confirmation_required: bool = False


class RiskAssessor:
    """Assesses risk levels for Excel operations."""
    
    def __init__(self):
        """Initialize the risk assessor with predefined risk rules."""
        self._operation_risks = {
            # Safe operations (read-only or minimal impact)
            'query_data': RiskLevel.LOW,
            'filter_data': RiskLevel.LOW,
            'aggregate_data': RiskLevel.LOW,
            'sort_data': RiskLevel.LOW,
            'create_chart': RiskLevel.LOW,
            'modify_chart': RiskLevel.LOW,
            'shift_axis': RiskLevel.LOW,
            'transform_values': RiskLevel.LOW,
            'resize_chart': RiskLevel.LOW,
            'insert_row': RiskLevel.LOW,
            
            # Medium risk operations (structural changes)
            'insert_column': RiskLevel.MEDIUM,
            'update_cells': RiskLevel.MEDIUM,
            
            # High risk operations (destructive)
            'delete_rows': RiskLevel.HIGH,
            'delete_columns': RiskLevel.HIGH,
            'clear_data': RiskLevel.HIGH,
            
            # Dangerous operations (mass operations)
            'format_all': RiskLevel.DANGEROUS,
            'delete_all': RiskLevel.DANGEROUS,
            'clear_all': RiskLevel.DANGEROUS,
            'replace_all': RiskLevel.DANGEROUS,
        }
        
        self._dangerous_keywords = [
            'all', 'everything', 'entire', 'whole', 'complete',
            'format all', 'delete all', 'clear all', 'remove all',
            'entire spreadsheet', 'whole sheet', 'complete file'
        ]
        
        self._high_risk_keywords = [
            'delete', 'remove', 'clear', 'drop', 'truncate',
            'format', 'restructure', 'reorganize'
        ]
    
    def assess_operation(self, operation: str, parameters: Dict[str, Any], 
                        command_text: str = "") -> RiskAssessment:
        """
        Assess the risk level of an operation.
        
        Args:
            operation: The operation name
            parameters: Operation parameters
            command_text: Original natural language command
            
        Returns:
            RiskAssessment with level, score, and reasons
        """
        reasons = []
        base_risk = self._operation_risks.get(operation, RiskLevel.MEDIUM)
        risk_score = self._get_base_score(base_risk)
        
        # Check for dangerous keywords in command text
        if self._contains_dangerous_keywords(command_text.lower()):
            base_risk = RiskLevel.DANGEROUS
            risk_score = 1.0
            reasons.append("Contains dangerous mass operation keywords")
        
        # Analyze parameters for additional risk factors
        param_risk, param_reasons = self._assess_parameters(parameters, operation)
        risk_score = max(risk_score, param_risk)
        reasons.extend(param_reasons)
        
        # Determine final risk level based on score
        final_risk = self._score_to_risk_level(risk_score)
        
        # Determine if operation should be blocked or requires confirmation
        blocked = final_risk == RiskLevel.DANGEROUS
        
        # Special handling for read-only operations - they should never require confirmation
        read_only_operations = ['query_data', 'filter_data', 'aggregate_data', 'sort_data']
        if operation in read_only_operations:
            confirmation_required = False
        else:
            confirmation_required = final_risk in [RiskLevel.HIGH, RiskLevel.MEDIUM]
        
        return RiskAssessment(
            level=final_risk,
            score=risk_score,
            reasons=reasons,
            blocked=blocked,
            confirmation_required=confirmation_required
        )
    
    def _get_base_score(self, risk_level: RiskLevel) -> float:
        """Convert risk level to numeric score."""
        scores = {
            RiskLevel.LOW: 0.2,
            RiskLevel.MEDIUM: 0.5,
            RiskLevel.HIGH: 0.8,
            RiskLevel.DANGEROUS: 1.0
        }
        return scores[risk_level]
    
    def _contains_dangerous_keywords(self, text: str) -> bool:
        """Check if text contains dangerous operation keywords."""
        return any(keyword in text for keyword in self._dangerous_keywords)
    
    def _assess_parameters(self, parameters: Dict[str, Any], operation: str = "") -> tuple[float, List[str]]:
        """
        Assess risk based on operation parameters.
        
        Returns:
            Tuple of (risk_score, reasons)
        """
        risk_score = 0.0
        reasons = []
        
        # Check for mass operations based on scope
        if 'max_rows' in parameters:
            max_rows = parameters.get('max_rows', 0)
            if max_rows > 50:
                risk_score = max(risk_score, 1.0)
                reasons.append(f"Operation affects {max_rows} rows (limit: 50)")
            elif max_rows > 10:
                risk_score = max(risk_score, 0.8)
                reasons.append(f"Operation affects {max_rows} rows")
        
        # Check for range operations
        if 'range' in parameters:
            range_str = parameters.get('range', '')
            if self._is_large_range(range_str):
                risk_score = max(risk_score, 0.8)
                reasons.append("Operation affects large range of cells")
        
        # Check for conditions that might affect many rows (only for write operations)
        if 'conditions' in parameters:
            conditions = parameters.get('conditions', [])
            # Only consider lack of conditions dangerous for write operations
            write_operations = ['update_cells', 'delete_rows', 'delete_columns', 'clear_data']
            read_operations = ['query_data', 'filter_data', 'aggregate_data', 'sort_data']
            
            # Skip condition check for safe read operations
            if operation in read_operations:
                pass  # Read operations are safe even without conditions
            elif (not conditions or conditions == ['*'] or 'all' in str(conditions).lower()) and any(op in operation for op in write_operations):
                risk_score = max(risk_score, 1.0)
                reasons.append("Operation has no limiting conditions")
        
        # Check for format operations
        if any(param in parameters for param in ['format', 'style', 'formatting']):
            risk_score = max(risk_score, 0.6)
            reasons.append("Operation involves formatting changes")
        
        return risk_score, reasons
    
    def _is_large_range(self, range_str: str) -> bool:
        """Check if a range string represents a large range."""
        if not range_str:
            return False
        
        # Simple heuristic: check for entire column/row references
        large_range_patterns = [':', 'A:Z', '1:1000', 'entire', 'all']
        return any(pattern in range_str.lower() for pattern in large_range_patterns)
    
    def _score_to_risk_level(self, score: float) -> RiskLevel:
        """Convert numeric score back to risk level."""
        if score >= 1.0:
            return RiskLevel.DANGEROUS
        elif score >= 0.8:
            return RiskLevel.HIGH
        elif score >= 0.5:
            return RiskLevel.MEDIUM
        else:
            return RiskLevel.LOW
    
    def get_risk_explanation(self, assessment: RiskAssessment) -> str:
        """Get human-readable explanation of risk assessment."""
        explanations = {
            RiskLevel.LOW: "This is a safe operation with minimal risk to your data.",
            RiskLevel.MEDIUM: "This operation may modify your data structure. Review carefully.",
            RiskLevel.HIGH: "This operation could significantly impact your data. Confirmation required.",
            RiskLevel.DANGEROUS: "This operation is blocked for safety as it could cause data loss."
        }
        
        base_explanation = explanations[assessment.level]
        if assessment.reasons:
            reasons_text = "; ".join(assessment.reasons)
            return f"{base_explanation} Reasons: {reasons_text}"
        
        return base_explanation