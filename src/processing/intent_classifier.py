"""
Intent classification for natural language commands.

This module provides intent classification capabilities that work alongside
the LLM service to categorize and validate user commands.
"""

import re
import logging
from typing import Dict, List, Optional, Tuple, Any
from dataclasses import dataclass
from enum import Enum

from ..llm.ollama_service import LLMResponse


class IntentCategory(Enum):
    """Main intent categories for Excel operations."""
    DATA_OPERATIONS = "data_operations"
    CHART_OPERATIONS = "chart_operations"
    VISUALIZATION_OPERATIONS = "visualization_operations"
    QUERY_OPERATIONS = "query_operations"
    UNKNOWN = "unknown"


@dataclass
class IntentClassification:
    """Result of intent classification."""
    category: IntentCategory
    operation: str
    confidence: float
    keywords_matched: List[str]
    parameters_extracted: Dict[str, Any]
    suggestions: List[str]


class IntentClassifier:
    """
    Classifies user intents and extracts parameters from natural language commands.
    
    This class works as a fallback and validation layer for LLM-based intent detection,
    providing rule-based classification when LLM is unavailable or uncertain.
    """
    
    def __init__(self):
        """Initialize intent classifier with keyword patterns."""
        self.logger = logging.getLogger(__name__)
        
        # Define keyword patterns for each intent category
        self._intent_patterns = {
            IntentCategory.DATA_OPERATIONS: {
                "create_data": {
                    "keywords": ["add", "insert", "create", "new", "append"],
                    "patterns": [
                        r"add\s+(?:new\s+)?(?:row|entry|record|data)",
                        r"insert\s+(?:new\s+)?(?:row|entry|record|data)",
                        r"create\s+(?:new\s+)?(?:row|entry|record|data)"
                    ]
                },
                "query_data": {
                    "keywords": ["find", "search", "get", "show", "display", "list", "query"],
                    "patterns": [
                        r"find\s+(?:all\s+)?(?:rows|records|data)",
                        r"show\s+(?:me\s+)?(?:all\s+)?(?:rows|records|data)",
                        r"get\s+(?:all\s+)?(?:rows|records|data)",
                        r"list\s+(?:all\s+)?(?:rows|records|data)"
                    ]
                },
                "update_data": {
                    "keywords": ["update", "modify", "change", "edit", "set"],
                    "patterns": [
                        r"update\s+(?:the\s+)?(?:row|record|data|value)",
                        r"change\s+(?:the\s+)?(?:row|record|data|value)",
                        r"modify\s+(?:the\s+)?(?:row|record|data|value)"
                    ]
                },
                "delete_data": {
                    "keywords": ["delete", "remove", "clear", "drop"],
                    "patterns": [
                        r"delete\s+(?:the\s+)?(?:row|record|data)",
                        r"remove\s+(?:the\s+)?(?:row|record|data)",
                        r"clear\s+(?:the\s+)?(?:row|record|data)"
                    ]
                }
            },
            IntentCategory.CHART_OPERATIONS: {
                "shift_axis": {
                    "keywords": ["shift", "move", "offset", "translate"],
                    "patterns": [
                        r"shift\s+(?:the\s+)?(?:chart|graph|plot)\s+(?:left|right|up|down)",
                        r"move\s+(?:the\s+)?(?:chart|graph|plot)\s+(?:left|right|up|down)",
                        r"(?:left|right|up|down)\s+shift"
                    ]
                },
                "transform_values": {
                    "keywords": ["transform", "modify", "change", "apply", "subtract", "add", "multiply", "divide"],
                    "patterns": [
                        r"(?:subtract|add|multiply|divide)\s+\d+",
                        r"(?:reduce|increase)\s+(?:all\s+)?(?:values|data)",
                        r"transform\s+(?:the\s+)?(?:values|data)"
                    ]
                }
            },
            IntentCategory.VISUALIZATION_OPERATIONS: {
                "create_chart": {
                    "keywords": ["chart", "plot", "graph", "visualize", "create"],
                    "patterns": [
                        r"create\s+(?:a\s+)?(?:bar|line|pie|scatter|area)\s+chart",
                        r"(?:make|generate)\s+(?:a\s+)?(?:chart|plot|graph)",
                        r"visualize\s+(?:the\s+)?data",
                        r"plot\s+(?:the\s+)?data"
                    ]
                }
            }
        }
        
        # Common parameter extraction patterns
        self._parameter_patterns = {
            "numbers": r"\b\d+(?:\.\d+)?\b",
            "sheet_names": r"sheet\s+['\"]?([^'\"]+)['\"]?",
            "column_names": r"column\s+['\"]?([^'\"]+)['\"]?",
            "ranges": r"[A-Z]+\d+:[A-Z]+\d+",
            "directions": r"\b(left|right|up|down|x|y)\b",
            "chart_types": r"\b(bar|line|pie|scatter|area|doughnut|radar)\b",
            "operations": r"\b(add|subtract|multiply|divide|increase|decrease|reduce)\b"
        }
    
    def classify_intent(self, user_command: str, llm_response: Optional[LLMResponse] = None) -> IntentClassification:
        """
        Classify user intent from natural language command.
        
        Args:
            user_command: Natural language command from user
            llm_response: Optional LLM response for validation/enhancement
            
        Returns:
            IntentClassification: Classification result with confidence score
        """
        command_lower = user_command.lower().strip()
        
        # If LLM response is available, validate and enhance it
        if llm_response and llm_response.intent != "clarification_needed":
            return self._validate_llm_classification(command_lower, llm_response)
        
        # Perform rule-based classification
        return self._rule_based_classification(command_lower)
    
    def _validate_llm_classification(self, command: str, llm_response: LLMResponse) -> IntentClassification:
        """Validate and enhance LLM classification with rule-based checks."""
        # Convert LLM intent to our category system
        category = self._map_llm_intent_to_category(llm_response.intent)
        
        # Extract additional parameters using rule-based patterns
        rule_parameters = self._extract_parameters(command)
        
        # Merge LLM parameters with rule-based parameters
        merged_parameters = {**llm_response.parameters, **rule_parameters}
        
        # Find matching keywords for confidence calculation
        keywords_matched = self._find_matching_keywords(command, category, llm_response.operation)
        
        # Calculate confidence based on keyword matches and LLM confidence
        confidence = self._calculate_confidence(keywords_matched, llm_response.confidence)
        
        return IntentClassification(
            category=category,
            operation=llm_response.operation,
            confidence=confidence,
            keywords_matched=keywords_matched,
            parameters_extracted=merged_parameters,
            suggestions=[]
        )
    
    def _rule_based_classification(self, command: str) -> IntentClassification:
        """Perform rule-based intent classification."""
        best_match = None
        best_confidence = 0.0
        best_keywords = []
        
        # Check each intent category
        for category, operations in self._intent_patterns.items():
            for operation, config in operations.items():
                # Check keyword matches
                keywords_matched = []
                for keyword in config["keywords"]:
                    if keyword in command:
                        keywords_matched.append(keyword)
                
                # Check pattern matches
                pattern_matches = 0
                for pattern in config["patterns"]:
                    if re.search(pattern, command, re.IGNORECASE):
                        pattern_matches += 1
                
                # Calculate confidence score
                keyword_score = len(keywords_matched) / len(config["keywords"])
                pattern_score = min(pattern_matches, 1)  # Cap at 1.0
                confidence = (keyword_score * 0.6) + (pattern_score * 0.4)
                
                # Update best match if this is better
                if confidence > best_confidence:
                    best_confidence = confidence
                    best_match = (category, operation)
                    best_keywords = keywords_matched
        
        # Extract parameters
        parameters = self._extract_parameters(command)
        
        # Generate suggestions if confidence is low
        suggestions = []
        if best_confidence < 0.3:
            suggestions = self._generate_suggestions(command)
        
        if best_match:
            category, operation = best_match
        else:
            category, operation = IntentCategory.UNKNOWN, "unknown"
        
        return IntentClassification(
            category=category,
            operation=operation,
            confidence=best_confidence,
            keywords_matched=best_keywords,
            parameters_extracted=parameters,
            suggestions=suggestions
        )
    
    def _extract_parameters(self, command: str) -> Dict[str, Any]:
        """Extract parameters from command using regex patterns."""
        parameters = {}
        
        # Extract numbers
        numbers = re.findall(self._parameter_patterns["numbers"], command)
        if numbers:
            parameters["numbers"] = [float(n) for n in numbers]
        
        # Extract sheet names
        sheet_match = re.search(self._parameter_patterns["sheet_names"], command, re.IGNORECASE)
        if sheet_match:
            parameters["sheet_name"] = sheet_match.group(1)
        
        # Extract column names
        column_match = re.search(self._parameter_patterns["column_names"], command, re.IGNORECASE)
        if column_match:
            parameters["column_name"] = column_match.group(1)
        
        # Extract ranges
        range_matches = re.findall(self._parameter_patterns["ranges"], command, re.IGNORECASE)
        if range_matches:
            parameters["data_range"] = range_matches[0]
        
        # Extract directions
        direction_matches = re.findall(self._parameter_patterns["directions"], command, re.IGNORECASE)
        if direction_matches:
            parameters["direction"] = direction_matches[0].lower()
        
        # Extract chart types
        chart_type_matches = re.findall(self._parameter_patterns["chart_types"], command, re.IGNORECASE)
        if chart_type_matches:
            parameters["chart_type"] = chart_type_matches[0].lower()
        
        # Extract mathematical operations
        operation_matches = re.findall(self._parameter_patterns["operations"], command, re.IGNORECASE)
        if operation_matches:
            parameters["math_operation"] = operation_matches[0].lower()
        
        return parameters
    
    def _map_llm_intent_to_category(self, llm_intent: str) -> IntentCategory:
        """Map LLM intent string to our IntentCategory enum."""
        intent_mapping = {
            "data_operations": IntentCategory.DATA_OPERATIONS,
            "chart_operations": IntentCategory.CHART_OPERATIONS,
            "visualization_operations": IntentCategory.VISUALIZATION_OPERATIONS,
            "query_operations": IntentCategory.QUERY_OPERATIONS
        }
        
        return intent_mapping.get(llm_intent, IntentCategory.UNKNOWN)
    
    def _find_matching_keywords(self, command: str, category: IntentCategory, operation: str) -> List[str]:
        """Find keywords that match the given category and operation."""
        if category not in self._intent_patterns:
            return []
        
        if operation not in self._intent_patterns[category]:
            return []
        
        keywords = self._intent_patterns[category][operation]["keywords"]
        matched = []
        
        for keyword in keywords:
            if keyword in command.lower():
                matched.append(keyword)
        
        return matched
    
    def _calculate_confidence(self, keywords_matched: List[str], llm_confidence: float = 0.0) -> float:
        """Calculate overall confidence score."""
        keyword_confidence = min(len(keywords_matched) * 0.2, 1.0)  # Max 1.0
        
        if llm_confidence > 0:
            # Combine LLM confidence with keyword confidence
            return (llm_confidence * 0.7) + (keyword_confidence * 0.3)
        else:
            # Use only keyword confidence
            return keyword_confidence
    
    def _generate_suggestions(self, command: str) -> List[str]:
        """Generate suggestions for unclear commands."""
        suggestions = []
        
        # Check for common ambiguities
        if "chart" in command.lower() or "plot" in command.lower():
            suggestions.append("Try specifying the chart type (bar, line, pie, etc.)")
            suggestions.append("Include the data range you want to visualize")
        
        if "data" in command.lower():
            suggestions.append("Specify which sheet contains the data")
            suggestions.append("Be more specific about what you want to do with the data")
        
        if not suggestions:
            suggestions.append("Try being more specific about what you want to do")
            suggestions.append("Include details like sheet names, column names, or data ranges")
        
        return suggestions
    
    def get_intent_examples(self, category: IntentCategory) -> Dict[str, List[str]]:
        """Get example commands for a specific intent category."""
        examples = {
            IntentCategory.DATA_OPERATIONS: [
                "Add a new row with data: John, 25, Engineer",
                "Find all rows where status is active",
                "Update the price to 100 where product is laptop",
                "Delete rows where quantity is 0"
            ],
            IntentCategory.CHART_OPERATIONS: [
                "Shift the chart left by 2 units",
                "Move the graph right by 0.5",
                "Subtract 5 from all Y values",
                "Multiply all chart data by 2"
            ],
            IntentCategory.VISUALIZATION_OPERATIONS: [
                "Create a bar chart from sales data",
                "Make a line plot of monthly revenue",
                "Generate a pie chart for category distribution",
                "Visualize the data in columns A1:C10"
            ]
        }
        
        return {
            "examples": examples.get(category, []),
            "operations": list(self._intent_patterns.get(category, {}).keys()) if category in self._intent_patterns else []
        }
    
    def validate_parameters(self, classification: IntentClassification, required_params: List[str]) -> Tuple[bool, List[str]]:
        """
        Validate that required parameters are present in the classification.
        
        Args:
            classification: Intent classification result
            required_params: List of required parameter names
            
        Returns:
            Tuple of (is_valid, missing_parameters)
        """
        missing_params = []
        
        for param in required_params:
            if param not in classification.parameters_extracted:
                missing_params.append(param)
        
        return len(missing_params) == 0, missing_params