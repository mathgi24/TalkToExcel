"""Ollama LLM service for natural language command processing."""

import json
import time
import requests
from typing import Dict, Any, Optional, List
from dataclasses import dataclass
from pathlib import Path

from config.config_manager import config

# Add src directory to Python path for imports
import sys
from pathlib import Path
current_dir = Path(__file__).parent
src_dir = current_dir.parent if current_dir.name != 'src' else current_dir
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))



@dataclass
class LLMResponse:
    """Structured response from LLM service."""
    intent: str
    operation: str
    parameters: Dict[str, Any]
    confirmation_required: bool
    risk_assessment: str
    confidence: float = 0.0
    raw_response: str = ""


class OllamaConnectionError(Exception):
    """Raised when connection to Ollama fails."""
    pass


class OllamaService:
    """Service for interacting with Ollama LLM for command processing."""
    
    def __init__(self, template_registry=None):
        """Initialize Ollama service with configuration."""
        self.config = config.get_ollama_config()
        self.endpoint = self.config.get('endpoint', 'http://localhost:11434')
        self.model = self.config.get('model', 'mistral:7b-instruct')
        self.temperature = self.config.get('temperature', 0.1)
        self.max_tokens = self.config.get('max_tokens', 1000)
        self.timeout = self.config.get('timeout', 30)
        self.retry_attempts = self.config.get('retry_attempts', 3)
        self.retry_delay = self.config.get('retry_delay', 2)
        
        self._session = requests.Session()
        
        # Use provided template_registry or create new one
        if template_registry is None:
            from templates.template_registry import TemplateRegistry
            self.template_registry = TemplateRegistry()
        else:
            self.template_registry = template_registry
        
        self._operation_templates = {}
        self._load_operation_templates()
    
    def _load_operation_templates(self) -> None:
        """Load operation templates dynamically from operations.yaml via template_registry."""
        try:
            # Load all operations with examples from the template registry
            self._operation_templates = self.template_registry.get_all_operations_with_examples()
            
            # Log successful loading
            stats = self.template_registry.get_registry_stats()
            print(f"✅ Loaded {stats['total_operations']} operations from {stats['categories']} categories for LLM service")
            
        except Exception as e:
            print(f"❌ Failed to load operation templates: {str(e)}")
            # Fallback to minimal hardcoded templates
            self._operation_templates = {
                "data_operations": {
                    "query_data": {
                        "parameters": ["sheet_name", "conditions", "columns"],
                        "safety_level": "safe",
                        "intent_keywords": ["find", "search", "query", "get", "show"],
                        "examples": ["show me the data", "find all rows"]
                    }
                }
            }
    
    def initialize_connection(self) -> bool:
        """Initialize and test connection to Ollama with enhanced error handling.
        
        Returns:
            True if connection successful, False otherwise
            
        Raises:
            OllamaConnectionError: If connection fails after retries
        """
        last_error = None
        
        for attempt in range(self.retry_attempts):
            try:
                # Test basic connectivity first
                response = self._session.get(
                    f"{self.endpoint}/api/tags",
                    timeout=self.timeout
                )
                response.raise_for_status()
                
                # Check if our model is available
                models = response.json().get('models', [])
                model_names = [model.get('name', '') for model in models]
                
                if self.model not in model_names:
                    # Provide specific error for missing model
                    available_models = ', '.join(model_names) if model_names else 'None'
                    raise OllamaConnectionError(
                        f"Model '{self.model}' not found. Available models: {available_models}. "
                        f"Download with: ollama pull {self.model}"
                    )
                
                # Test model functionality with a simple request
                test_payload = {
                    "model": self.model,
                    "prompt": "Test connection",
                    "stream": False,
                    "options": {"num_predict": 1}
                }
                
                test_response = self._session.post(
                    f"{self.endpoint}/api/generate",
                    json=test_payload,
                    timeout=10  # Shorter timeout for test
                )
                test_response.raise_for_status()
                
                return True
                
            except requests.exceptions.ConnectionError as e:
                last_error = OllamaConnectionError(
                    f"Cannot connect to Ollama at {self.endpoint}. "
                    f"Ensure Ollama is running with: ollama serve"
                )
                
            except requests.exceptions.Timeout as e:
                last_error = OllamaConnectionError(
                    f"Connection to Ollama timed out after {self.timeout}s. "
                    f"Check if Ollama is responding or increase timeout."
                )
                
            except requests.exceptions.HTTPError as e:
                if e.response.status_code == 404:
                    last_error = OllamaConnectionError(
                        f"Ollama API endpoint not found. Check if Ollama is running at {self.endpoint}"
                    )
                else:
                    last_error = OllamaConnectionError(
                        f"HTTP error {e.response.status_code}: {e.response.text}"
                    )
                    
            except requests.exceptions.RequestException as e:
                last_error = OllamaConnectionError(f"Request failed: {str(e)}")
                
            except Exception as e:
                last_error = OllamaConnectionError(f"Unexpected error: {str(e)}")
            
            # Wait before retry with exponential backoff
            if attempt < self.retry_attempts - 1:
                wait_time = min(self.retry_delay * (2 ** attempt), 30)  # Cap at 30 seconds
                time.sleep(wait_time)
        
        # All attempts failed
        if last_error:
            raise last_error
        else:
            raise OllamaConnectionError("Connection failed after all retry attempts")
        
        return False
    
    def generate_system_prompt(self) -> str:
        """Generate dynamic system prompt based on available operations.
        
        Returns:
            System prompt string with all available operations
        """
        prompt = """You are an Excel operation classifier. Based on the user command, 
return JSON with intent and parameters for Excel operations.

Available operations:
"""
        
        for category, operations in self._operation_templates.items():
            prompt += f"\n{category.replace('_', ' ').title()}:\n"
            
            for op_name, op_config in operations.items():
                keywords = ', '.join(op_config['intent_keywords'])
                examples = '; '.join(op_config['examples'])
                parameters = ', '.join(op_config['parameters'])
                
                prompt += f"  - {op_name}: Keywords: {keywords}\n"
                prompt += f"    Examples: {examples}\n"
                prompt += f"    Parameters: {parameters}\n"
                prompt += f"    Safety: {op_config['safety_level']}\n\n"
        
        prompt += """
IMPORTANT: Return ONLY valid JSON in this exact format:
{
  "intent": "operation_category",
  "operation": "specific_operation_name",
  "parameters": {
    "param1": "value1",
    "param2": "value2"
  },
  "confirmation_required": true/false,
  "risk_assessment": "low/medium/high"
}

Rules:
1. Always return valid JSON
2. Use operation names exactly as listed above
3. Set confirmation_required=false for query_data operations (they are safe)
4. Set risk_assessment="low" for query_data operations
5. Include all relevant parameters from the user command
6. For "show/display/get" commands, use intent="data_operations" and operation="query_data"
7. For sheet names, if not specified, leave sheet_name empty (system will use default)
8. For "first N rows" commands, set limit=N in parameters
9. For "last N rows" commands, set limit=N and sort_order="desc" in parameters
10. If command is unclear, set intent="clarification_needed"
11. For sheet names, use full names like "Sales Data", "Employees", not abbreviated forms
12. For conditions, always use dictionary format: {"column_name": "value"}, never use strings
13. If user mentions a sheet name partially (like "sales" for "Sales Data"), use the full name
14. ALWAYS include "limit": 100 for query_data operations unless user specifies a different number
15. For create_data operations with key=value format (like "Name=John, Age=25"), parse into data object: {"Name": "John", "Age": 25}
16. For create_data operations with space-separated values (like "John Sales 50000"), parse into data array: ["John", "Sales", 50000]
17. Convert string numbers to actual numbers (100000 → 100000, not "100000")
18. Convert string booleans to actual booleans (True → true, False → false)
19. For create_data operations, always set confirmation_required=true and risk_assessment="low"

Examples:
- "show me the first 5 rows" → {"intent": "data_operations", "operation": "query_data", "parameters": {"limit": 5}, "confirmation_required": false, "risk_assessment": "low"}
- "display data from Sales Data sheet" → {"intent": "data_operations", "operation": "query_data", "parameters": {"sheet_name": "Sales Data", "limit": 100}, "confirmation_required": false, "risk_assessment": "low"}
- "show employee rows" → {"intent": "data_operations", "operation": "query_data", "parameters": {"sheet_name": "Employees", "limit": 100}, "confirmation_required": false, "risk_assessment": "low"}
- "show first 2 rows in sales data" → {"intent": "data_operations", "operation": "query_data", "parameters": {"sheet_name": "Sales Data", "limit": 2}, "confirmation_required": false, "risk_assessment": "low"}
- "show inventory with stock = 200" → {"intent": "data_operations", "operation": "query_data", "parameters": {"sheet_name": "Inventory", "conditions": {"Stock": 200}, "limit": 100}, "confirmation_required": false, "risk_assessment": "low"}
- "show active employees" → {"intent": "data_operations", "operation": "query_data", "parameters": {"sheet_name": "Employees", "conditions": {"Active": true}, "limit": 100}, "confirmation_required": false, "risk_assessment": "low"}
- "show employees with salary > 50000" → {"intent": "data_operations", "operation": "query_data", "parameters": {"sheet_name": "Employees", "conditions": {"Salary": {"operator": ">", "value": 50000}}, "limit": 100}, "confirmation_required": false, "risk_assessment": "low"}
- "add new row with Product=Laptop, Price=1200" → {"intent": "data_operations", "operation": "create_data", "parameters": {"sheet_name": "Products", "data": {"Product": "Laptop", "Price": 1200}}, "confirmation_required": true, "risk_assessment": "low"}
- "add new row to Employees sheet with Name=Ram, Department=AI, Salary=100000, Active=True" → {"intent": "data_operations", "operation": "create_data", "parameters": {"sheet_name": "Employees", "data": {"Name": "Ram", "Department": "AI", "Salary": 100000, "Active": true}}, "confirmation_required": true, "risk_assessment": "low"}
- "insert employee with Name=John, Department=Sales, Salary=50000" → {"intent": "data_operations", "operation": "create_data", "parameters": {"sheet_name": "Employees", "data": {"Name": "John", "Department": "Sales", "Salary": 50000}}, "confirmation_required": true, "risk_assessment": "low"}
- "create new entry with Product=Phone, Quantity=50, Price=800" → {"intent": "data_operations", "operation": "create_data", "parameters": {"data": {"Product": "Phone", "Quantity": 50, "Price": 800}}, "confirmation_required": true, "risk_assessment": "low"}
- "add row Ram AI 100000 True" → {"intent": "data_operations", "operation": "create_data", "parameters": {"data": ["Ram", "AI", 100000, true]}, "confirmation_required": true, "risk_assessment": "low"}
- "create bar chart from sales data" → {"intent": "visualization_operations", "operation": "create_chart", "parameters": {"sheet_name": "Sales Data", "data_range": "Sales Data", "chart_type": "bar"}, "confirmation_required": false, "risk_assessment": "low"}
- "create chart from inventory data" → {"intent": "visualization_operations", "operation": "create_chart", "parameters": {"sheet_name": "Inventory", "data_range": "Inventory", "chart_type": "bar"}, "confirmation_required": false, "risk_assessment": "low"}
- "make a pie chart from employee data" → {"intent": "visualization_operations", "operation": "create_chart", "parameters": {"sheet_name": "Employees", "data_range": "Employees", "chart_type": "pie"}, "confirmation_required": false, "risk_assessment": "low"}
"""
        
        return prompt
    
    def parse_to_structured_command(self, user_command: str) -> LLMResponse:
        """Convert natural language command to structured JSON with enhanced error handling.
        
        Args:
            user_command: Natural language command from user
            
        Returns:
            LLMResponse with structured command data
            
        Raises:
            OllamaConnectionError: If LLM request fails
        """
        system_prompt = self.generate_system_prompt()
        
        payload = {
            "model": self.model,
            "prompt": f"System: {system_prompt}\n\nUser: {user_command}\n\nAssistant:",
            "stream": False,
            "options": {
                "temperature": self.temperature,
                "num_predict": self.max_tokens
            }
        }
        
        last_error = None
        
        for attempt in range(self.retry_attempts):
            try:
                response = self._session.post(
                    f"{self.endpoint}/api/generate",
                    json=payload,
                    timeout=self.timeout
                )
                response.raise_for_status()
                
                result = response.json()
                raw_response = result.get('response', '').strip()
                
                # Validate response is not empty
                if not raw_response:
                    raise OllamaConnectionError("Empty response from LLM")
                
                # Parse JSON response
                try:
                    parsed_response = json.loads(raw_response)
                    return self._create_llm_response(parsed_response, raw_response)
                    
                except json.JSONDecodeError:
                    # Try to extract JSON from response if it's wrapped in text
                    json_start = raw_response.find('{')
                    json_end = raw_response.rfind('}') + 1
                    
                    if json_start >= 0 and json_end > json_start:
                        json_str = raw_response[json_start:json_end]
                        try:
                            parsed_response = json.loads(json_str)
                            return self._create_llm_response(parsed_response, raw_response)
                        except json.JSONDecodeError:
                            pass
                    
                    # If JSON parsing fails, return clarification needed
                    return LLMResponse(
                        intent="clarification_needed",
                        operation="parse_error",
                        parameters={
                            "error": "Failed to parse LLM response", 
                            "raw": raw_response[:500],  # Truncate for logging
                            "suggestion": "Try rephrasing your command more clearly"
                        },
                        confirmation_required=True,
                        risk_assessment="high",
                        raw_response=raw_response
                    )
                
            except requests.exceptions.ConnectionError as e:
                last_error = OllamaConnectionError(
                    f"Connection lost to Ollama (attempt {attempt + 1}/{self.retry_attempts}). "
                    f"Check if Ollama is still running."
                )
                
            except requests.exceptions.Timeout as e:
                last_error = OllamaConnectionError(
                    f"Request timed out after {self.timeout}s (attempt {attempt + 1}/{self.retry_attempts}). "
                    f"The model may be processing a complex request."
                )
                
            except requests.exceptions.HTTPError as e:
                if e.response.status_code == 404:
                    last_error = OllamaConnectionError(
                        f"Model '{self.model}' not found. Ensure it's downloaded: ollama pull {self.model}"
                    )
                elif e.response.status_code == 500:
                    last_error = OllamaConnectionError(
                        f"Ollama server error. Try restarting Ollama service."
                    )
                else:
                    last_error = OllamaConnectionError(
                        f"HTTP {e.response.status_code}: {e.response.text}"
                    )
                
            except requests.exceptions.RequestException as e:
                last_error = OllamaConnectionError(f"Request failed: {str(e)}")
                
            except Exception as e:
                last_error = OllamaConnectionError(f"Unexpected error during LLM request: {str(e)}")
            
            # Wait before retry with exponential backoff
            if attempt < self.retry_attempts - 1:
                wait_time = min(self.retry_delay * (2 ** attempt), 30)
                time.sleep(wait_time)
        
        # All attempts failed
        if last_error:
            raise last_error
        else:
            raise OllamaConnectionError("LLM request failed after all retry attempts")
    
    def _create_llm_response(self, parsed_data: Dict[str, Any], raw_response: str) -> LLMResponse:
        """Create LLMResponse from parsed JSON data.
        
        Args:
            parsed_data: Parsed JSON response from LLM
            raw_response: Raw response string
            
        Returns:
            LLMResponse object
        """
        return LLMResponse(
            intent=parsed_data.get('intent', 'unknown'),
            operation=parsed_data.get('operation', 'unknown'),
            parameters=parsed_data.get('parameters', {}),
            confirmation_required=parsed_data.get('confirmation_required', True),
            risk_assessment=parsed_data.get('risk_assessment', 'medium'),
            confidence=parsed_data.get('confidence', 0.0),
            raw_response=raw_response
        )
    
    def assess_command_safety(self, command: str) -> str:
        """Assess the safety level of a command.
        
        Args:
            command: Natural language command
            
        Returns:
            Safety level: 'low', 'medium', or 'high'
        """
        dangerous_keywords = [
            'delete all', 'remove all', 'clear all', 'format all',
            'delete everything', 'remove everything', 'clear everything'
        ]
        
        medium_risk_keywords = [
            'delete', 'remove', 'clear', 'update', 'modify',
            'change', 'replace'
        ]
        
        command_lower = command.lower()
        
        # Check for dangerous commands first (more specific patterns)
        for keyword in dangerous_keywords:
            if keyword in command_lower:
                return 'high'
        
        # Check for medium risk commands
        for keyword in medium_risk_keywords:
            if keyword in command_lower:
                return 'medium'
        
        return 'low'
    
    def generate_confirmation_prompt(self, operation: Dict[str, Any]) -> str:
        """Generate user confirmation prompt for operations.
        
        Args:
            operation: Operation dictionary with intent, operation, and parameters
            
        Returns:
            Confirmation prompt string
        """
        intent = operation.get('intent', 'unknown')
        op_name = operation.get('operation', 'unknown')
        params = operation.get('parameters', {})
        risk = operation.get('risk_assessment', 'medium')
        
        prompt = f"Confirm {intent} operation: {op_name}\n"
        prompt += f"Risk level: {risk.upper()}\n"
        
        if params:
            prompt += "Parameters:\n"
            for key, value in params.items():
                prompt += f"  - {key}: {value}\n"
        
        prompt += "\nProceed with this operation? (y/n): "
        
        return prompt
    
    def validate_response(self, response: LLMResponse) -> bool:
        """Validate LLM response structure and content.
        
        Args:
            response: LLMResponse to validate
            
        Returns:
            True if response is valid, False otherwise
        """
        # Check required fields
        if not response.intent or not response.operation:
            return False
        
        # Check if operation exists in templates
        operation_found = False
        for category, operations in self._operation_templates.items():
            if response.operation in operations:
                operation_found = True
                break
        
        if not operation_found and response.intent != "clarification_needed":
            return False
        
        # Check risk assessment values
        valid_risk_levels = ['low', 'medium', 'high']
        if response.risk_assessment not in valid_risk_levels:
            return False
        
        return True
    
    def get_available_operations(self) -> Dict[str, List[str]]:
        """Get list of available operations by category.
        
        Returns:
            Dictionary mapping categories to operation lists
        """
        result = {}
        for category, operations in self._operation_templates.items():
            result[category] = list(operations.keys())
        return result
    
    def is_available(self) -> bool:
        """Check if the LLM service is available.
        
        Returns:
            True if service is available, False otherwise
        """
        try:
            return self.initialize_connection()
        except Exception:
            return False
    
    def cleanup(self):
        """Clean up resources used by the service."""
        if hasattr(self, '_session'):
            self._session.close()