"""
Calculator Plugin
Allows SAM to perform exact mathematical calculations instead of guessing.
"""

import math
import operator
import re

PLUGIN = {
    "name": "calculator",
    "description": (
        "Use this tool to calculate exact mathematical equations and arithmetic. "
        "Do NOT guess math answers. Always use this tool for multiplication, division, addition, etc."
    ),
    "parameters": {
        "type": "OBJECT",
        "properties": {
            "expression": {
                "type": "STRING", 
                "description": "The math expression to evaluate (e.g., '847 * 293', '100 / 3', '2^8')"
            },
        },
        "required": ["expression"],
    },
}

def run(parameters: dict, player=None, session_memory=None) -> str:
    expr = parameters.get("expression", "")
    
    if not expr:
        return "No expression provided."

    if player:
        player.write_log(f"SAM: Calculating '{expr}'...")

    # Basic sanitization
    expr = expr.replace('^', '**')
    expr = expr.replace('x', '*')
    
    # Remove any characters that aren't numbers, operators, or spaces
    allowed_chars = set("0123456789+-*/(). **")
    if not all(c in allowed_chars for c in expr):
        return "Error: Expression contains invalid characters."

    try:
        # Evaluate safely
        result = eval(expr, {"__builtins__": None}, {})
        
        # Format large numbers nicely
        if isinstance(result, (int, float)):
            formatted_result = f"{result:,}"
        else:
            formatted_result = str(result)
            
        return f"The exact answer is {formatted_result}."
    except Exception as e:
        return f"Sorry, I couldn't calculate that: {str(e)}"
