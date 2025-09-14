"""Dynamic prompt generator that includes all available operations."""

from typing import Dict, Any, List, Optional
import json
from .template_loader import TemplateLoader
from .template_registry import TemplateRegistry


class PromptGenerator:
    """Generates dynamic LLM prompts based on available operations."""
    
    def __init__(self, template_loader: Optional[TemplateLoader] = None,
                 template_registry: Optional[TemplateRegistry] = None):
        """Initialize prompt generator.
        
        Args:
            template_loader: TemplateLoader instance
            template_registry: TemplateRegistry instance
        """
        self.template_loader = template_loader or TemplateLoader()
        self.template_registry = template_registry or TemplateRegistry(self.template_loader)
    
    def generate_system_prompt(self) -> str:
        """Generate the main system prompt for the LLM.
        
        Returns:
            Complete system prompt string
        """
        operations = self.template_loader.get_operations()
        
        prompt = """You are an Excel operation classifier and command interpreter. Your role is to analyze natural language commands and convert them into structured JSON operations that can be executed on Excel files.

## Your Task
Convert user commands into structured JSON format that specifies:
1. The operation intent (what type of operation to perform)
2. The specific operation name
3. Required parameters for execution
4. Safety assessment and confirmation requirements

## Available Operations

"""
        
        # Add operation categories and details
        for category_name, category_ops in operations.items():
            category_title = category_name.replace('_', ' ').title()
            prompt += f"### {category_title}\n\n"
            
            for op_name, op_config in category_ops.items():
                prompt += f"**{op_name}**\n"
                prompt += f"- Description: {op_config['description']}\n"
                prompt += f"- Parameters: {', '.join(op_config['parameters'])}\n"
                prompt += f"- Safety Level: {op_config['safety_level']}\n"
                prompt += f"- Keywords: {', '.join(op_config['intent_keywords'])}\n"
                prompt += f"- Examples:\n"
                for example in op_config['examples']:
                    prompt += f"  - \"{example}\"\n"
                prompt += "\n"
        
        prompt += """## Response Format

You MUST respond with valid JSON in this exact format:

```json
{
  "intent": "operation_category",
  "operation": "specific_operation_name",
  "parameters": {
    "param1": "value1",
    "param2": "value2"
  },
  "confidence": 0.95,
  "safety_level": "safe|medium|high",
  "confirmation_required": true|false,
  "reasoning": "Brief explanation of why this operation was chosen"
}
```

## Parameter Guidelines

1. **Extract specific values** from user commands when possible
2. **Use reasonable defaults** for missing parameters
3. **Request clarification** by setting confidence < 0.7 when ambiguous
4. **Assess safety level** based on operation impact:
   - safe: Read operations, single cell updates, simple queries
   - medium: Multiple row operations, chart modifications, formula changes  
   - high: Delete operations, format operations, structural changes

## Safety Rules

1. **Block dangerous operations** that could affect entire spreadsheets
2. **Require confirmation** for any delete or destructive operations
3. **Limit scope** of operations (max 50 rows per operation)
4. **Validate parameters** to prevent unintended consequences

## Examples

User: "shift chart left by 2"
Response:
```json
{
  "intent": "chart_operations",
  "operation": "shift_axis", 
  "parameters": {
    "chart_id": "auto_detect",
    "axis": "x",
    "amount": -2
  },
  "confidence": 0.95,
  "safety_level": "safe",
  "confirmation_required": false,
  "reasoning": "Clear chart axis shift command with specific direction and amount"
}
```

User: "delete all rows where status is inactive"
Response:
```json
{
  "intent": "data_operations",
  "operation": "delete_rows",
  "parameters": {
    "sheet_name": "auto_detect",
    "conditions": {"status": "inactive"},
    "max_rows": 50
  },
  "confidence": 0.90,
  "safety_level": "high", 
  "confirmation_required": true,
  "reasoning": "Destructive delete operation requires confirmation and scope limitation"
}
```

## Important Notes

- Always respond with valid JSON only
- Never execute operations directly - only classify and structure them
- When uncertain, lower confidence score and request clarification
- Consider the safety implications of each operation
- Preserve user intent while ensuring safe execution
"""
        
        return prompt
    
    def generate_operation_summary(self) -> str:
        """Generate a summary of all available operations.
        
        Returns:
            Formatted summary string
        """
        operations = self.template_loader.get_operations()
        registry_stats = self.template_registry.get_registry_stats()
        
        summary = "# Available Operations Summary\n\n"
        summary += f"**Total Operations:** {registry_stats['total_operations']}\n"
        summary += f"**Implemented:** {registry_stats['implemented_operations']}\n"
        summary += f"**Categories:** {registry_stats['categories']}\n\n"
        
        for category_name in registry_stats['category_names']:
            category_ops = operations.get(category_name, {})
            category_title = category_name.replace('_', ' ').title()
            summary += f"## {category_title} ({len(category_ops)} operations)\n\n"
            
            for op_name, op_config in category_ops.items():
                implemented = self.template_registry.is_operation_available(category_name, op_name)
                status = "✅" if implemented else "⚠️"
                summary += f"- {status} **{op_name}**: {op_config['description']}\n"
            
            summary += "\n"
        
        return summary
    
    def generate_category_prompt(self, category: str) -> str:
        """Generate prompt for a specific operation category.
        
        Args:
            category: Operation category name
            
        Returns:
            Category-specific prompt
            
        Raises:
            KeyError: If category does not exist
        """
        try:
            category_ops = self.template_loader.get_operations_by_category(category)
        except KeyError:
            raise KeyError(f"Category '{category}' not found")
        
        category_title = category.replace('_', ' ').title()
        prompt = f"# {category_title} Operations\n\n"
        prompt += f"You are specialized in {category_title.lower()} operations for Excel files.\n\n"
        prompt += "## Available Operations\n\n"
        
        for op_name, op_config in category_ops.items():
            prompt += f"### {op_name}\n"
            prompt += f"{op_config['description']}\n\n"
            prompt += f"**Parameters:** {', '.join(op_config['parameters'])}\n"
            prompt += f"**Safety Level:** {op_config['safety_level']}\n"
            prompt += f"**Keywords:** {', '.join(op_config['intent_keywords'])}\n\n"
            prompt += "**Examples:**\n"
            for example in op_config['examples']:
                prompt += f"- {example}\n"
            prompt += "\n"
        
        return prompt
    
    def generate_safety_prompt(self, safety_level: str) -> str:
        """Generate prompt for operations of a specific safety level.
        
        Args:
            safety_level: Safety level ('safe', 'medium', 'high')
            
        Returns:
            Safety-level specific prompt
        """
        operations = self.template_loader.get_operations_by_safety_level(safety_level)
        
        prompt = f"# {safety_level.title()} Operations\n\n"
        
        if safety_level == 'safe':
            prompt += "These operations are safe to execute without confirmation:\n\n"
        elif safety_level == 'medium':
            prompt += "These operations require careful parameter validation:\n\n"
        else:  # high
            prompt += "These operations are potentially destructive and require confirmation:\n\n"
        
        for category_name, category_ops in operations.items():
            category_title = category_name.replace('_', ' ').title()
            prompt += f"## {category_title}\n\n"
            
            for op_name, op_config in category_ops.items():
                prompt += f"- **{op_name}**: {op_config['description']}\n"
        
        return prompt
    
    def generate_examples_prompt(self) -> str:
        """Generate prompt with all operation examples.
        
        Returns:
            Examples-focused prompt
        """
        operations = self.template_loader.get_operations()
        
        prompt = "# Operation Examples\n\n"
        prompt += "Here are examples of natural language commands and their corresponding operations:\n\n"
        
        for category_name, category_ops in operations.items():
            category_title = category_name.replace('_', ' ').title()
            prompt += f"## {category_title}\n\n"
            
            for op_name, op_config in category_ops.items():
                prompt += f"### {op_name}\n"
                for example in op_config['examples']:
                    prompt += f"- \"{example}\"\n"
                prompt += "\n"
        
        return prompt
    
    def generate_validation_prompt(self, user_command: str) -> str:
        """Generate prompt for validating a specific user command.
        
        Args:
            user_command: User's natural language command
            
        Returns:
            Validation-focused prompt
        """
        prompt = f"""Analyze this user command and determine the best matching operation:

Command: "{user_command}"

Consider:
1. Which operation category best matches the intent?
2. What specific operation should be performed?
3. What parameters can be extracted from the command?
4. What is the safety level of this operation?
5. Are there any ambiguities that need clarification?

Available operations:
"""
        
        # Add condensed operation list
        operations = self.template_loader.get_operations()
        for category_name, category_ops in operations.items():
            prompt += f"\n{category_name}:\n"
            for op_name, op_config in category_ops.items():
                keywords = ', '.join(op_config['intent_keywords'][:3])  # First 3 keywords
                prompt += f"  - {op_name}: {keywords}\n"
        
        return prompt
    
    def get_operation_keywords_map(self) -> Dict[str, List[str]]:
        """Get mapping of keywords to operation identifiers.
        
        Returns:
            Dictionary mapping keywords to list of 'category.operation' strings
        """
        keyword_map = {}
        operations = self.template_loader.get_operations()
        
        for category_name, category_ops in operations.items():
            for op_name, op_config in category_ops.items():
                operation_id = f"{category_name}.{op_name}"
                
                for keyword in op_config.get('intent_keywords', []):
                    keyword_lower = keyword.lower()
                    if keyword_lower not in keyword_map:
                        keyword_map[keyword_lower] = []
                    keyword_map[keyword_lower].append(operation_id)
        
        return keyword_map
    
    def suggest_operations_for_command(self, user_command: str) -> List[Dict[str, Any]]:
        """Suggest possible operations for a user command based on keywords.
        
        Args:
            user_command: User's natural language command
            
        Returns:
            List of suggested operations with confidence scores
        """
        command_lower = user_command.lower()
        keyword_map = self.get_operation_keywords_map()
        suggestions = []
        
        # Score operations based on keyword matches
        operation_scores = {}
        for keyword, operation_ids in keyword_map.items():
            if keyword in command_lower:
                for op_id in operation_ids:
                    if op_id not in operation_scores:
                        operation_scores[op_id] = 0
                    operation_scores[op_id] += 1
        
        # Convert to suggestions with metadata
        operations = self.template_loader.get_operations()
        for op_id, score in operation_scores.items():
            category, operation = op_id.split('.', 1)
            op_config = operations[category][operation]
            
            suggestions.append({
                'category': category,
                'operation': operation,
                'score': score,
                'confidence': min(score / 3.0, 1.0),  # Normalize to 0-1
                'description': op_config['description'],
                'safety_level': op_config['safety_level']
            })
        
        # Sort by score descending
        suggestions.sort(key=lambda x: x['score'], reverse=True)
        
        return suggestions[:5]  # Return top 5 suggestions