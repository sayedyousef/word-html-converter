# ============= utils.py =============
"""Utility functions"""
import zipfile
from pathlib import Path
from lxml import etree
import hashlib

def get_file_hash(filepath):
    """Get MD5 hash of file for caching"""
    with open(filepath, 'rb') as f:
        return hashlib.md5(f.read()).hexdigest()

def extract_xml_from_docx(docx_path, xml_path):
    """Extract specific XML file from docx"""
    with zipfile.ZipFile(docx_path, 'r') as z:
        if xml_path in z.namelist():
            return z.read(xml_path)
    return None

def clean_latex_string(latex):
    """Clean and normalize LaTeX string"""
    if not latex:
        return ""
    
    # Remove unnecessary whitespace
    latex = latex.strip()
    
    # Ensure math delimiters
    if not latex.startswith('$'):
        latex = f"${latex}$"
    
    return latex

def create_equation_anchor(equation_id, latex_text):
    """Create HTML anchor for equation"""
    # Simple anchor format that can be styled later
    return f'<span class="equation" id="eq_{equation_id}">{latex_text}</span>'

