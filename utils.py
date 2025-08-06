# utils.py
"""Utility functions for document processing."""

import re
import unicodedata
from pathlib import Path
from typing import Optional, Tuple
from typing import Optional, Tuple, List

def sanitize_filename(filename: str) -> str:
    """Sanitize filename for filesystem compatibility."""
    # Remove invalid characters
    filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
    
    # Normalize unicode characters
    filename = unicodedata.normalize('NFKD', filename)
    
    # Limit length
    if len(filename) > 200:
        filename = filename[:200]
    
    return filename.strip()

def extract_text_safely(paragraph) -> str:
    """Safely extract text from paragraph handling encoding issues."""
    try:
        return paragraph.text or ""
    except Exception as e:
        # Handle any encoding issues
        return ""

def detect_latex_equations(text: str) -> Tuple[bool, List[str]]:
    """Detect LaTeX equations in text."""
    equations = []
    
    # Check for common LaTeX patterns
    latex_patterns = [
        r'\$[^$]+\$',                    # Inline math $...$
        r'\$\$[^$]+\$\$',                # Display math $$...$$
        r'\\\[[^\]]+\\\]',               # Display \[...\]  <- FIXED
        r'\\\([^)]+\\\)',                # Inline \(...\)   <- FIXED
        r'\\begin\{equation\}.*?\\end\{equation\}',  # Equation environment
        r'\\begin\{align\}.*?\\end\{align\}',        # Align environment
    ]
    
    for pattern in latex_patterns:
        matches = re.findall(pattern, text, re.DOTALL)
        equations.extend(matches)
    
    # Also check for common LaTeX commands
    latex_commands = [
        r'\\frac', r'\\sqrt', r'\\sum', r'\\int', r'\\alpha', r'\\beta',
        r'\\gamma', r'\\delta', r'\\partial', r'\\infty', r'\\pm'
    ]
    
    has_latex = any(cmd in text for cmd in latex_commands) or len(equations) > 0
    
    return has_latex, equations

def format_article_number(index: int) -> str:
    """Format article number with leading zeros."""
    return f"{index:03d}_"

