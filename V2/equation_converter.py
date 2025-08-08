# ============= equation_converter.py =============
"""Core equation extraction and conversion logic"""
import zipfile
from pathlib import Path
from lxml import etree
import re
from typing import List, Dict, Tuple
from logger import setup_logger
from utils import extract_xml_from_docx, clean_latex_string, create_equation_anchor

logger = setup_logger("equation_converter")

class EquationConverter:
    """Handle equation extraction and conversion from Word documents"""
    
    # OMML namespace
    MATH_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
    WORD_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    
    def __init__(self, docx_path):
        self.docx_path = Path(docx_path)
        self.equations = []
        self.equation_map = {}  # Maps equation ID to LaTeX
        
    def extract_equations(self) -> List[Dict]:
        """Extract all equations from document - OMML, LaTeX, and other formats"""
        logger.info(f"Extracting equations from {self.docx_path.name}")
        
        # Get document.xml
        doc_xml = extract_xml_from_docx(self.docx_path, 'word/document.xml')
        if not doc_xml:
            logger.warning("No document.xml found")
            return []
        
        # Parse XML
        root = etree.fromstring(doc_xml)
        
        equations = []
        
        # 1. Find OMML math elements (native Word equations)
        omml_equations = self._find_omml_equations(root)
        equations.extend(omml_equations)
        
        # 2. Find LaTeX equations ($ ... $ or \[ ... \] in text)
        latex_equations = self._find_latex_equations(root)
        equations.extend(latex_equations)
        
        # 3. Find MathType or field-based equations
        field_equations = self._find_field_equations(root)
        equations.extend(field_equations)
        
        # 4. Check for inline equations in paragraphs
        inline_equations = self._find_inline_equations(root)
        equations.extend(inline_equations)
        
        logger.info(f"Found {len(equations)} equations: "
                    f"{len(omml_equations)} OMML, "
                    f"{len(latex_equations)} LaTeX, "
                    f"{len(field_equations)} field-based")
        
        self.equations = equations
        return equations
    
    def _find_omml_equations(self, root) -> List[Dict]:
        """Find OMML equation blocks"""
        equations = []
        
        # Use namespace prefix in xpath, not the full namespace URI
        # Define namespace map for xpath
        namespaces = {
            'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        }
        
        # Look for oMathPara and oMath elements using namespace prefix
        for i, math_elem in enumerate(root.xpath('//m:oMathPara | //m:oMath', namespaces=namespaces)):
            eq_data = {
                'id': f"block_{i}",
                'type': 'block',
                'omml': etree.tostring(math_elem, encoding='unicode'),
                'position': i
            }
            equations.append(eq_data)
            
        return equations

    def _find_inline_equations(self, root) -> List[Dict]:
        """Find inline equations in paragraphs"""
        equations = []
        
        # Define namespace map
        namespaces = {
            'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        }
        
        # Look for equations embedded in paragraphs
        paragraphs = root.xpath('//w:p', namespaces=namespaces)
        
        for p_idx, para in enumerate(paragraphs):
            # Check for math elements within paragraph
            math_elems = para.xpath('.//m:oMath', namespaces=namespaces)
            
            for m_idx, math_elem in enumerate(math_elems):
                eq_data = {
                    'id': f"inline_{p_idx}_{m_idx}",
                    'type': 'inline',
                    'omml': etree.tostring(math_elem, encoding='unicode'),
                    'position': p_idx
                }
                equations.append(eq_data)
                
        return equations

    def _find_latex_equations(self, root) -> List[Dict]:
        """Find LaTeX equations in document text"""
        equations = []
        latex_patterns = [
            (r'\$\$(.+?)\$\$', 'display'),  # Display math $$...$$
            (r'\$(.+?)\$', 'inline'),       # Inline math $...$
            (r'\\\[(.+?)\\\]', 'display'),  # Display \[...\]
            (r'\\\((.+?)\\\)', 'inline'),   # Inline \(...\)
        ]
        
        # Get all text runs
        text_runs = root.xpath('//w:t', namespaces={'w': self.WORD_NS})
        
        for idx, text_elem in enumerate(text_runs):
            if text_elem.text:
                text = text_elem.text
                
                # Check for LaTeX patterns
                for pattern, eq_type in latex_patterns:
                    matches = re.finditer(pattern, text)
                    for match in matches:
                        latex_content = match.group(1)
                        equations.append({
                            'id': f"latex_{idx}_{match.start()}",
                            'type': 'latex',
                            'subtype': eq_type,
                            'content': latex_content,
                            'raw': match.group(0),
                            'position': idx,
                            'element': text_elem
                        })
        
        return equations

    def _find_field_equations(self, root) -> List[Dict]:
        """Find equations in Word fields (EQ fields, MathType)"""
        equations = []
        
        # Look for field codes
        fields = root.xpath('//w:instrText', namespaces={'w': self.WORD_NS})
        
        for idx, field in enumerate(fields):
            if field.text:
                text = field.text.strip()
                
                # Check for EQ field (Word's old equation format)
                if text.startswith('EQ '):
                    equations.append({
                        'id': f"field_eq_{idx}",
                        'type': 'field',
                        'subtype': 'eq_field',
                        'content': text[3:].strip(),
                        'position': idx,
                        'element': field
                    })
                
                # Check for embedded LaTeX in fields
                elif 'latex' in text.lower() or '\\' in text:
                    equations.append({
                        'id': f"field_latex_{idx}",
                        'type': 'field',
                        'subtype': 'latex_field',
                        'content': text,
                        'position': idx,
                        'element': field
                    })
        
        return equations

    def convert_to_latex(self, equation_data: Dict) -> str:
        """Convert equation to LaTeX based on its type"""
        logger.debug(f"Converting equation {equation_data['id']} of type {equation_data['type']}")
        
        eq_type = equation_data['type']
        
        # LaTeX equations are already in LaTeX format
        if eq_type == 'latex':
            latex = equation_data['content']
            # Just clean it up
            return clean_latex_string(latex)
        
        # OMML equations need conversion
        elif eq_type == 'omml':
            return self._convert_omml_to_latex(equation_data)
        
        # Field equations might be LaTeX or need parsing
        elif eq_type == 'field':
            if equation_data['subtype'] == 'latex_field':
                # Already LaTeX
                return clean_latex_string(equation_data['content'])
            else:
                # EQ field needs special parsing
                return self._convert_eq_field(equation_data['content'])
        
        # Inline equations in mixed format
        elif eq_type == 'inline':
            # Try to detect format
            content = equation_data.get('content', '')
            if '$' in content or '\\' in content:
                # Likely LaTeX
                return clean_latex_string(content)
            else:
                # Try OMML conversion
                return self._convert_omml_to_latex(equation_data)
        
        # Default fallback
        return "\\text{[unknown equation type]}"

    def _convert_omml_to_latex(self, equation_data: Dict) -> str:
        """Convert OMML equation to LaTeX (existing method renamed)"""
        omml = equation_data.get('omml', equation_data.get('content', ''))
        
        # Try dwml first
        latex = self._convert_with_dwml(omml)
        if latex:
            return latex
        
        # Fallback to basic parsing
        return self._convert_basic(omml)

    def _convert_eq_field(self, field_content: str) -> str:
        """Convert Word EQ field to LaTeX"""
        # EQ fields use special syntax like EQ \f(1,2) for fractions
        content = field_content.strip()
        
        # Common EQ field patterns
        replacements = {
            r'\\f\(([^,]+),([^)]+)\)': r'\\frac{\1}{\2}',  # Fractions
            r'\\r\(([^)]+)\)': r'\\sqrt{\1}',              # Square root
            r'\\s\(([^,]+),([^)]+)\)': r'\1^{\2}',         # Superscript
            r'\\i\(([^,]+),([^,]+),([^)]+)\)': r'\\int_{\1}^{\2} \3',  # Integral
        }
        
        for pattern, replacement in replacements.items():
            content = re.sub(pattern, replacement, content)
        
        return clean_latex_string(content)
    
    def _convert_with_dwml(self, omml_xml: str) -> str:
        """Convert using dwml library"""
        try:
            from dwml import omml
            # Parse OMML and convert
            equations = omml.loads(omml_xml)
            if equations:
                return equations[0].latex
        except ImportError:
            logger.debug("dwml not available")
        except Exception as e:
            logger.debug(f"dwml conversion failed: {e}")
        return None
    
    def _convert_with_docxlatex(self, omml_xml: str) -> str:
        """Convert using docxlatex library"""
        try:
            # This would need custom implementation since docxlatex
            # works on whole documents, not XML fragments
            pass
        except:
            pass
        return None
    
    def _convert_basic(self, omml_xml: str) -> str:
        """Basic conversion using pattern matching"""
        logger.debug("Using basic pattern matching for conversion")
        
        # Parse the OMML
        try:
            root = etree.fromstring(omml_xml)
        except:
            # If it's a fragment, wrap it
            omml_xml = f'<root xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">{omml_xml}</root>'
            root = etree.fromstring(omml_xml)
        
        # Extract text content (very basic)
        text_parts = []
        
        # Look for common elements
        for elem in root.iter():
            # Get text from run elements
            if elem.tag.endswith('}t'):
                if elem.text:
                    text_parts.append(elem.text)
        
        # Join and create basic LaTeX
        if text_parts:
            latex = ' '.join(text_parts)
            # Try to detect if it's a fraction, superscript, etc.
            latex = self._apply_basic_patterns(latex)
            return latex
        
        return "\\text{[equation]}"
    
    def _apply_basic_patterns(self, text: str) -> str:
        """Apply basic LaTeX patterns"""
        # Simple replacements
        replacements = {
            '^2': '^{2}',
            '^3': '^{3}',
            'sqrt': '\\sqrt',
            'alpha': '\\alpha',
            'beta': '\\beta',
            'gamma': '\\gamma',
            'pi': '\\pi',
        }
        
        for old, new in replacements.items():
            text = text.replace(old, new)
        
        return text
    
    def replace_equations_with_anchors(self, content: str) -> str:
        """Replace equations in content with anchor text"""
        logger.info("Replacing equations with LaTeX anchors")
        
        modified_content = content
        
        for eq_id, latex in self.equation_map.items():
            # Create anchor
            anchor = create_equation_anchor(eq_id, latex)
            
            # For now, append at end (in real implementation, 
            # would replace at actual position)
            modified_content += f"\n{anchor}\n"
        
        return modified_content

