# ============= equation_replacer_inplace.py =============
"""Replace equations IN-PLACE without reconstructing document"""
import zipfile
from pathlib import Path
from lxml import etree
import shutil
from logger import setup_logger

logger = setup_logger("equation_replacer")

class InPlaceEquationReplacer:
    """Modify Word document IN-PLACE - preserve everything except equations"""
    
    def __init__(self, docx_path):
        self.docx_path = Path(docx_path)
        self.equations_found = []
        
    def replace_equations_in_place(self, output_path=None):
        """Replace ONLY equations, keep everything else identical"""
        
        if not output_path:
            output_path = self.docx_path.parent / f"{self.docx_path.stem}_equations_as_text.docx"
        
        # Copy original document
        shutil.copy2(self.docx_path, output_path)
        
        # Open as ZIP and modify ONLY document.xml
        temp_zip = output_path.with_suffix('.tmp')
        
        with zipfile.ZipFile(output_path, 'r') as zin:
            with zipfile.ZipFile(temp_zip, 'w', zipfile.ZIP_DEFLATED) as zout:
                # Copy all files
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    
                    # Only modify document.xml
                    if item.filename == 'word/document.xml':
                        data = self._replace_equations_in_xml(data)
                    
                    zout.writestr(item, data)
        
        # Replace original with modified
        temp_zip.replace(output_path)
        
        logger.info(f"Modified document saved to {output_path}")
        logger.info(f"Replaced {len(self.equations_found)} equations with LaTeX text")
        
        return output_path
    
    def _replace_equations_in_xml(self, doc_xml_bytes):
        """Replace OMML equations with text runs containing LaTeX"""
        
        root = etree.fromstring(doc_xml_bytes)
        
        # Namespaces
        ns = {
            'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'xml': 'http://www.w3.org/XML/1998/namespace'
        }
        
        # Find all math elements
        math_elements = root.xpath('//m:oMath', namespaces=ns)
        
        for idx, math_elem in enumerate(math_elements):
            # Extract LaTeX from OMML
            latex = self._extract_latex_from_omml(math_elem, ns)
            
            # Create unique bookmark/anchor ID
            eq_id = f"eq_{idx}_{hash(latex)&0xFFFF:04x}"
            
            # Create replacement structure with bookmark
            replacement = self._create_text_with_bookmark(latex, eq_id, ns)
            
            # Replace math element with text + bookmark
            parent = math_elem.getparent()
            if parent is not None:
                # Insert replacement at same position
                index = list(parent).index(math_elem)
                parent.remove(math_elem)
                for i, elem in enumerate(replacement):
                    parent.insert(index + i, elem)
            
            # Store equation info
            self.equations_found.append({
                'id': eq_id,
                'latex': latex,
                'position': idx
            })
        
        return etree.tostring(root, encoding='utf-8', xml_declaration=True)
    
    def _extract_latex_from_omml(self, math_elem, ns):
        """Extract LaTeX representation from OMML"""
        
        # Collect all text elements
        texts = []
        
        # Look for specific structures
        # Check for subscripts (sSub)
        for ssub in math_elem.xpath('.//m:sSub', namespaces=ns):
            base = self._get_text_from_element(ssub.find('.//m:e', ns), ns)
            sub = self._get_text_from_element(ssub.find('.//m:sub', ns), ns)
            texts.append(f"{base}_{{{sub}}}")
        
        # Check for superscripts (sSup)
        for ssup in math_elem.xpath('.//m:sSup', namespaces=ns):
            base = self._get_text_from_element(ssup.find('.//m:e', ns), ns)
            sup = self._get_text_from_element(ssup.find('.//m:sup', ns), ns)
            texts.append(f"{base}^{{{sup}}}")
        
        # Check for fractions
        for frac in math_elem.xpath('.//m:f', namespaces=ns):
            num = self._get_text_from_element(frac.find('.//m:num', ns), ns)
            den = self._get_text_from_element(frac.find('.//m:den', ns), ns)
            texts.append(f"\\frac{{{num}}}{{{den}}}")
        
        # If no specific structures found, get all text
        if not texts:
            for t_elem in math_elem.xpath('.//m:t', namespaces=ns):
                if t_elem.text:
                    text = t_elem.text
                    # Handle special symbols
                    text = text.replace('≠', '\\neq')
                    text = text.replace('≤', '\\leq')
                    text = text.replace('≥', '\\geq')
                    text = text.replace('∞', '\\infty')
                    texts.append(text)
        
        return ' '.join(texts) if texts else "[equation]"
    
    def _get_text_from_element(self, elem, ns):
        """Get all text from an element"""
        if elem is None:
            return ""
        
        texts = []
        for t in elem.xpath('.//m:t', namespaces=ns):
            if t.text:
                texts.append(t.text)
        return ''.join(texts)
    
    def _create_text_with_bookmark(self, latex, eq_id, ns):
        """Create Word text run with bookmark for equation"""
        
        elements = []
        
        # Create bookmark start
        bookmark_start = etree.Element(f"{{{ns['w']}}}bookmarkStart")
        bookmark_start.set(f"{{{ns['w']}}}id", str(hash(eq_id) & 0xFFFF))
        bookmark_start.set(f"{{{ns['w']}}}name", eq_id)
        elements.append(bookmark_start)
        
        # Create text run with LaTeX
        run = etree.Element(f"{{{ns['w']}}}r")
        
        # Add run properties for formatting
        rPr = etree.SubElement(run, f"{{{ns['w']}}}rPr")
        
        # Make it look like equation (monospace, background)
        highlight = etree.SubElement(rPr, f"{{{ns['w']}}}highlight")
        highlight.set(f"{{{ns['w']}}}val", "lightGray")
        
        rFonts = etree.SubElement(rPr, f"{{{ns['w']}}}rFonts")
        rFonts.set(f"{{{ns['w']}}}ascii", "Courier New")
        rFonts.set(f"{{{ns['w']}}}hAnsi", "Courier New")
        
        # Add text
        text = etree.SubElement(run, f"{{{ns['w']}}}t")
        text.set(f"{{{ns['xml']}}}space", "preserve")
        text.text = f" {latex} "
        
        elements.append(run)
        
        # Create bookmark end
        bookmark_end = etree.Element(f"{{{ns['w']}}}bookmarkEnd")
        bookmark_end.set(f"{{{ns['w']}}}id", str(hash(eq_id) & 0xFFFF))
        elements.append(bookmark_end)
        
        return elements
    
    def save_equation_mapping(self, output_path):
        """Save equation mapping to separate file"""
        import json
        
        mapping_file = output_path.parent / f"{output_path.stem}_equations.json"
        
        with open(mapping_file, 'w', encoding='utf-8') as f:
            json.dump({
                'source': str(self.docx_path),
                'output': str(output_path),
                'equations': self.equations_found
            }, f, indent=2, ensure_ascii=False)
        
        logger.info(f"Equation mapping saved to {mapping_file}")

# ============= For HTML: Use mammoth but post-process =============
class MammothWithEquationFix:
    """Use mammoth for HTML but fix equations after"""
    
    def __init__(self, docx_path):
        self.docx_path = Path(docx_path)
        
    def convert_to_html_with_equations(self, output_path=None):
        """Convert to HTML and fix equation display"""
        import mammoth
        
        if not output_path:
            output_path = self.docx_path.parent / f"{self.docx_path.stem}.html"
        
        # First, replace equations in Word doc
        replacer = InPlaceEquationReplacer(self.docx_path)
        temp_docx = self.docx_path.parent / f"{self.docx_path.stem}_temp.docx"
        replacer.replace_equations_in_place(temp_docx)
        
        # Now convert modified doc to HTML with mammoth
        with open(temp_docx, "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file)
            html = result.value
        
        # Fix LaTeX in HTML (remove escape sequences)
        html = self._fix_latex_in_html(html)
        
        # Add CSS for equations
        css = """
        <style>
            .equation {
                background-color: #f5f5f5;
                padding: 2px 6px;
                font-family: 'Courier New', monospace;
                border-radius: 3px;
                display: inline-block;
                margin: 2px;
            }
        </style>
        """
        html = css + html
        
        # Save HTML
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html)
        
        # Clean up temp file
        temp_docx.unlink()
        
        logger.info(f"HTML saved to {output_path}")
        return output_path
    
    def _fix_latex_in_html(self, html):
        """Fix LaTeX formatting in HTML"""
        import re
        
        # Remove backslash escapes that mammoth adds
        html = html.replace(r'\\', '\\')
        
        # Wrap equations in spans
        # Pattern for LaTeX-like content
        pattern = r'(\$[^$]+\$|\\[a-zA-Z]+\{[^}]*\}|\w+_\{[^}]*\}|\w+\^\{[^}]*\})'
        html = re.sub(pattern, r'<span class="equation">\1</span>', html)
        
        return html

# ============= Updated main processing =============
def process_document_preserving_structure(docx_path, output_format='both'):
    """Process document preserving original structure"""
    
    docx_path = Path(docx_path)
    
    if output_format in ['docx', 'both']:
        # Replace equations in Word doc
        replacer = InPlaceEquationReplacer(docx_path)
        docx_output = docx_path.parent / f"{docx_path.stem}_equations_text.docx"
        replacer.replace_equations_in_place(docx_output)
        replacer.save_equation_mapping(docx_output)
        print(f"✓ Word document saved: {docx_output}")
    
    if output_format in ['html', 'both']:
        # Convert to HTML with fixed equations
        converter = MammothWithEquationFix(docx_path)
        html_output = docx_path.parent / f"{docx_path.stem}.html"
        converter.convert_to_html_with_equations(html_output)
        print(f"✓ HTML saved: {html_output}")

if __name__ == "__main__":
    # Test with your file
    test_file = Path(r"D:\Work 3 (20-Oct-24)\2 Side projects May 25\Encyclopedia\articles\مقالات بعد الاخراج\test\الدالة واحد لواحد (جاهزة للنشر).docx")
    process_document_preserving_structure(test_file, 'both')