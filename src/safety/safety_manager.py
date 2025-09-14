"""
Safety manager that coordinates all safety mechanisms.

Provides a unified interface for risk assessment, scope analysis,
command blocking, and parameter validation.
"""

from typing import Dict, List, Any, Optional
from dataclasses import dataclass

from .risk_assessor import RiskAssessor, RiskAssessment, RiskLevel
from .scope_analyzer import ScopeAnalyzer, ScopeAnalysis
from .command_blocker import CommandBlocker, BlockResult
from .parameter_validator import ParameterValidator, ValidationResult


@dataclass
class SafetyResult:
    """Comprehensive safety analysis result."""
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


class SafetyManager:
    """Coordinates all safety mechanisms for Excel operations."""
    
    def __init__(self, max_rows: int = 50, max_columns: int = 20):
        """
        Initialize safety manager with all safety components.
        
        Args:
            max_rows: Maximum rows that can be affected by an operation
            max_columns: Maximum columns that can be affected by an operation
        """
        self.risk_assessor = RiskAssessor()
        self.scope_analyzer = ScopeAnalyzer(max_rows, max_columns)
        self.command_blocker = CommandBlocker()
        self.parameter_validator = ParameterValidator()
        
        self.max_rows = max_rows
        self.max_columns = max_columns
    
    def evaluate_operation(self, operation: str, parameters: Dict[str, Any],
                          command_text: str = "", 
                          sheet_info: Optional[Dict[str, Any]] = None) -> SafetyResult:
        """
        Perform comprehensive safety evaluation of an operation.
        
        Args:
            operation: The operation name
            parameters: Operation parameters
            command_text: Original natural language command
            sheet_info: Information about the target sheet
            
        Returns:
            SafetyResult with comprehensive safety analysis
        """
        error_messages = []
        warnings = []
        suggestions = []
        
        # 1. Validate parameters first
        validation_result = self.parameter_validator.validate_parameters(operation, parameters)
        if not validation_result.valid:
            error_messages.extend(validation_result.errors)
        warnings.extend(validation_result.warnings)
        
        # Use sanitized parameters for further analysis
        safe_parameters = validation_result.sanitized_parameters
        
        # 2. Check if command should be blocked
        block_result = self.command_blocker.check_command(operation, safe_parameters, command_text)
        if block_result.blocked:
            error_messages.append(block_result.error_message)
            suggestions.extend(block_result.suggested_alternatives)
        
        # 3. Assess risk level
        risk_assessment = self.risk_assessor.assess_operation(operation, safe_parameters, command_text)
        if risk_assessment.blocked:
            error_messages.append(f"Operation blocked due to {risk_assessment.level.value} risk")
        
        # 4. Analyze operation scope
        scope_analysis = self.scope_analyzer.analyze_scope(operation, safe_parameters, sheet_info)
        if not scope_analysis.within_limits:
            error_messages.extend(scope_analysis.violations)
            suggestions.extend(scope_analysis.suggested_alternatives)
        
        # Determine overall safety status
        safe = (validation_result.valid and 
                not block_result.blocked and 
                not risk_assessment.blocked and 
                scope_analysis.within_limits)
        
        blocked = (block_result.blocked or 
                  risk_assessment.blocked or 
                  not scope_analysis.within_limits)
        
        confirmation_required = (risk_assessment.confirmation_required and 
                               not blocked)
        
        # Collect additional suggestions
        if risk_assessment.level in [RiskLevel.MEDIUM, RiskLevel.HIGH]:
            suggestions.append("Consider creating a backup before proceeding")
        
        if scope_analysis.estimated_rows > 10:
            suggestions.append("Large operations may take longer to complete")
        
        return SafetyResult(
            safe=safe,
            blocked=blocked,
            confirmation_required=confirmation_required,
            risk_assessment=risk_assessment,
            scope_analysis=scope_analysis,
            block_result=block_result,
            validation_result=validation_result,
            error_messages=error_messages,
            warnings=warnings,
            suggestions=suggestions
        )
    
    def get_safety_summary(self, safety_result: SafetyResult) -> str:
        """Get a human-readable summary of the safety analysis."""
        if safety_result.blocked:
            return f"❌ Operation blocked: {'; '.join(safety_result.error_messages)}"
        
        if not safety_result.safe:
            return f"⚠️ Operation has safety concerns: {'; '.join(safety_result.error_messages)}"
        
        if safety_result.confirmation_required:
            risk_level = safety_result.risk_assessment.level.value
            scope = f"{safety_result.scope_analysis.estimated_rows} rows, {safety_result.scope_analysis.estimated_columns} columns"
            return f"⚠️ Confirmation required: {risk_level} risk operation affecting {scope}"
        
        return "✅ Operation is safe to proceed"
    
    def get_detailed_report(self, safety_result: SafetyResult) -> str:
        """Get a detailed safety report."""
        report = []
        
        # Header
        if safety_result.blocked:
            report.append("🚫 OPERATION BLOCKED")
        elif not safety_result.safe:
            report.append("⚠️ SAFETY CONCERNS DETECTED")
        elif safety_result.confirmation_required:
            report.append("❓ CONFIRMATION REQUIRED")
        else:
            report.append("✅ OPERATION APPROVED")
        
        report.append("=" * 50)
        
        # Risk Assessment
        report.append(f"Risk Level: {safety_result.risk_assessment.level.value.upper()}")
        report.append(f"Risk Score: {safety_result.risk_assessment.score:.2f}/1.0")
        if safety_result.risk_assessment.reasons:
            report.append("Risk Factors:")
            for reason in safety_result.risk_assessment.reasons:
                report.append(f"  • {reason}")
        
        # Scope Analysis
        report.append(f"\nScope: {safety_result.scope_analysis.estimated_rows} rows, {safety_result.scope_analysis.estimated_columns} columns")
        report.append(f"Within Limits: {'Yes' if safety_result.scope_analysis.within_limits else 'No'}")
        
        # Validation Results
        if safety_result.validation_result.errors:
            report.append("\nValidation Errors:")
            for error in safety_result.validation_result.errors:
                report.append(f"  • {error}")
        
        if safety_result.validation_result.warnings:
            report.append("\nValidation Warnings:")
            for warning in safety_result.validation_result.warnings:
                report.append(f"  • {warning}")
        
        # Error Messages
        if safety_result.error_messages:
            report.append("\nError Messages:")
            for error in safety_result.error_messages:
                report.append(f"  • {error}")
        
        # Suggestions
        if safety_result.suggestions:
            report.append("\nSuggestions:")
            for suggestion in safety_result.suggestions:
                report.append(f"  • {suggestion}")
        
        return "\n".join(report)
    
    def create_confirmation_prompt(self, safety_result: SafetyResult) -> str:
        """Create a confirmation prompt for operations requiring user approval."""
        if not safety_result.confirmation_required:
            return ""
        
        risk_level = safety_result.risk_assessment.level.value
        scope = f"{safety_result.scope_analysis.estimated_rows} rows and {safety_result.scope_analysis.estimated_columns} columns"
        
        prompt = f"""
⚠️ CONFIRMATION REQUIRED

This is a {risk_level} risk operation that will affect {scope}.

Risk factors:
"""
        
        for reason in safety_result.risk_assessment.reasons:
            prompt += f"• {reason}\n"
        
        if safety_result.warnings:
            prompt += "\nWarnings:\n"
            for warning in safety_result.warnings:
                prompt += f"• {warning}\n"
        
        prompt += "\nA backup will be created automatically before proceeding."
        prompt += "\n\nDo you want to continue? (yes/no): "
        
        return prompt
    
    def update_safety_limits(self, max_rows: int = None, max_columns: int = None):
        """Update safety limits for operations."""
        if max_rows is not None:
            self.max_rows = max_rows
            self.scope_analyzer.max_rows = max_rows
        
        if max_columns is not None:
            self.max_columns = max_columns
            self.scope_analyzer.max_columns = max_columns
    
    def add_custom_blocked_operation(self, operation: str, reason: str, 
                                   message: str, alternatives: List[str]):
        """Add a custom blocked operation."""
        self.command_blocker.add_blocked_operation(operation, reason, message, alternatives)
    
    def is_operation_safe(self, operation: str, parameters: Dict[str, Any],
                         command_text: str = "") -> bool:
        """Quick check if an operation is safe without detailed analysis."""
        safety_result = self.evaluate_operation(operation, parameters, command_text)
        return safety_result.safe and not safety_result.blocked
    
    def get_operation_requirements(self, operation: str) -> Dict[str, Any]:
        """Get safety requirements and parameter help for an operation."""
        param_help = self.parameter_validator.get_parameter_help(operation)
        
        return {
            'max_rows_limit': self.max_rows,
            'max_columns_limit': self.max_columns,
            'parameter_help': param_help,
            'blocked': self.command_blocker.is_operation_blocked(operation),
            'safe_alternatives': self.command_blocker.get_safe_alternatives(operation) if self.command_blocker.is_operation_blocked(operation) else []
        }