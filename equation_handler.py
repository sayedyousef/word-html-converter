# equation_handler.py
"""Handle LaTeX equation detection and processing."""

import re
import logging
from typing import Tuple, List

class EquationHandler:
    """Handles LaTeX equation processing."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def process_equations(self, text: str) -> str:
        """Process equations in text for proper HTML rendering."""
        # For now, we just ensure equations are properly delimited
        # MathJax/KaTeX will handle the rendering
        
        # Fix common issues
        text = self._fix_equation_delimiters(text)
        
        return text
    
    def _fix_equation_delimiters(self, text: str) -> str:
        """Fix common delimiter issues."""
        # Ensure proper spacing around delimiters
        text = re.sub(r'(?<![\\$])\$(?![\\$])', ' $ ', text)
        text = re.sub(r'\$\$', ' $$ ', text)
        
        # Clean up multiple spaces
        text = re.sub(r'\s+', ' ', text)
        
        return text

