import xml.etree.ElementTree as ET
from lxml import etree
import re

class OmmlToLatexConverter:
    """Simple OMML to LaTeX converter for common mathematical expressions"""
    
    def __init__(self):
        # Define OMML namespace
        self.ns = {'m': 'http://schemas.openxmlformats.org/officeDocument/2006/math'}
        
        # Mapping of common OMML elements to LaTeX
        self.symbol_map = {
            '∞': r'\infty',
            '±': r'\pm',
            '≤': r'\leq',
            '≥': r'\geq',
            '≠': r'\neq',
            '→': r'\rightarrow',
            '←': r'\leftarrow',
            '∈': r'\in',
            '∉': r'\notin',
            '∑': r'\sum',
            '∏': r'\prod',
            '∫': r'\int',
            '√': r'\sqrt',
            'α': r'\alpha',
            'β': r'\beta',
            'γ': r'\gamma',
            'δ': r'\delta',
            'π': r'\pi',
            'σ': r'\sigma',
            'θ': r'\theta',
            'λ': r'\lambda',
            'μ': r'\mu',
        }
    
    def convert_omml_to_latex(self, omml_string):
        """
        Convert OMML XML string to LaTeX
        
        Args:
            omml_string: OMML XML as string
            
        Returns:
            LaTeX string representation
        """
        try:
            # Parse OMML XML
            root = etree.fromstring(omml_string.encode('utf-8'))
            
            # Process the OMML tree
            latex = self._process_element(root)
            
            # Clean up the LaTeX
            latex = self._clean_latex(latex)
            
            return latex
        except Exception as e:
            return f"Error converting OMML: {str(e)}"
    
    def _process_element(self, element):
        """Recursively process OMML elements"""
        
        # Get the tag without namespace
        tag = element.tag.split('}')[-1] if '}' in element.tag else element.tag
        
        # Handle different OMML elements
        if tag == 'oMath' or tag == 'oMathPara':
            # Root math element
            return self._process_children(element)
        
        elif tag == 'r':
            # Run element (contains text)
            text = element.findtext('.//m:t', '', self.ns) or element.findtext('.//t', '')
            return self._convert_symbols(text)
        
        elif tag == 'frac':
            # Fraction
            num = self._process_element(element.find('.//m:num', self.ns))
            den = self._process_element(element.find('.//m:den', self.ns))
            return f'\\frac{{{num}}}{{{den}}}'
        
        elif tag == 'sup':
            # Superscript
            base = self._process_element(element.find('.//m:e', self.ns))
            sup = self._process_element(element.find('.//m:sup', self.ns))
            return f'{{{base}}}^{{{sup}}}'
        
        elif tag == 'sub':
            # Subscript
            base = self._process_element(element.find('.//m:e', self.ns))
            sub = self._process_element(element.find('.//m:sub', self.ns))
            return f'{{{base}}}_{{{sub}}}'
        
        elif tag == 'rad':
            # Radical (square root)
            deg = element.find('.//m:deg', self.ns)
            rad = self._process_element(element.find('.//m:e', self.ns))
            if deg is not None:
                degree = self._process_element(deg)
                return f'\\sqrt[{degree}]{{{rad}}}'
            return f'\\sqrt{{{rad}}}'
        
        elif tag == 'func':
            # Function
            fname = element.findtext('.//m:fName//m:t', '', self.ns)
            arg = self._process_element(element.find('.//m:e', self.ns))
            return f'\\{fname}({arg})'
        
        elif tag == 'd':
            # Delimiter (parentheses, brackets, etc.)
            begin = element.findtext('.//m:begChr', '(', self.ns)
            end = element.findtext('.//m:endChr', ')', self.ns)
            content = self._process_element(element.find('.//m:e', self.ns))
            
            # Convert delimiters to LaTeX
            begin = self._convert_delimiter(begin)
            end = self._convert_delimiter(end)
            
            return f'{begin}{content}{end}'
        
        elif tag == 'nary':
            # N-ary operator (sum, product, integral)
            op = element.findtext('.//m:chr', '', self.ns)
            sub = self._process_element(element.find('.//m:sub', self.ns))
            sup = self._process_element(element.find('.//m:sup', self.ns))
            e = self._process_element(element.find('.//m:e', self.ns))
            
            op_latex = self._convert_nary_operator(op)
            
            if sub and sup:
                return f'{op_latex}_{{{sub}}}^{{{sup}}} {e}'
            elif sub:
                return f'{op_latex}_{{{sub}}} {e}'
            else:
                return f'{op_latex} {e}'
        
        elif tag == 'acc':
            # Accent (hat, bar, dot, etc.)
            chr = element.findtext('.//m:accChr', '', self.ns)
            base = self._process_element(element.find('.//m:e', self.ns))
            
            accent_map = {
                '̂': 'hat',
                '̃': 'tilde',
                '̄': 'bar',
                '̇': 'dot',
                '̈': 'ddot',
                '⃗': 'vec',
            }
            
            accent = accent_map.get(chr, 'hat')
            return f'\\{accent}{{{base}}}'
        
        else:
            # Process children for unknown elements
            return self._process_children(element)
    
    def _process_children(self, element):
        """Process all child elements"""
        result = []
        for child in element:
            processed = self._process_element(child)
            if processed:
                result.append(processed)
        return ' '.join(result)
    
    def _convert_symbols(self, text):
        """Convert mathematical symbols to LaTeX"""
        for symbol, latex in self.symbol_map.items():
            text = text.replace(symbol, latex)
        return text
    
    def _convert_delimiter(self, delim):
        """Convert delimiters to LaTeX"""
        delimiter_map = {
            '[': '\\left[',
            ']': '\\right]',
            '{': '\\left\\{',
            '}': '\\right\\}',
            '(': '\\left(',
            ')': '\\right)',
            '|': '\\left|',
        }
        return delimiter_map.get(delim, delim)
    
    def _convert_nary_operator(self, op):
        """Convert n-ary operators to LaTeX"""
        nary_map = {
            '∑': '\\sum',
            '∏': '\\prod',
            '∫': '\\int',
            '∮': '\\oint',
            '⋃': '\\bigcup',
            '⋂': '\\bigcap',
        }
        return nary_map.get(op, op)
    
    def _clean_latex(self, latex):
        """Clean up the LaTeX output"""
        # Remove extra spaces
        latex = re.sub(r'\s+', ' ', latex)
        # Remove spaces around braces
        latex = re.sub(r'\s*{\s*', '{', latex)
        latex = re.sub(r'\s*}\s*', '}', latex)
        # Trim
        latex = latex.strip()
        return latex


def main():
    """Main function to demonstrate OMML to LaTeX conversion"""
    
    # Example OMML XML string (you would get this from your Word document)
    # This is a simple fraction example: 1/2
    sample_omml = '''
    <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <m:frac>
            <m:num>
                <m:r>
                    <m:t>1</m:t>
                </m:r>
            </m:num>
            <m:den>
                <m:r>
                    <m:t>2</m:t>
                </m:r>
            </m:den>
        </m:frac>
    </m:oMath>
    '''
    
    # Create converter instance
    converter = OmmlToLatexConverter()
    
    # Convert OMML to LaTeX
    latex_result = converter.convert_omml_to_latex(sample_omml)
    
    print("OMML Input (simplified view):")
    print("Fraction: 1/2")
    print("\nLaTeX Output:")
    print(latex_result)
    
    # Another example with superscript: x^2
    sample_omml2 = '''
    <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <m:sup>
            <m:e>
                <m:r>
                    <m:t>x</m:t>
                </m:r>
            </m:e>
            <m:sup>
                <m:r>
                    <m:t>2</m:t>
                </m:r>
            </m:sup>
        </m:sup>
    </m:oMath>
    '''
    
    latex_result2 = converter.convert_omml_to_latex(sample_omml2)
    print("\n" + "="*50)
    print("OMML Input (simplified view):")
    print("Superscript: x^2")
    print("\nLaTeX Output:")
    print(latex_result2)
    
    # Example with square root: √(a+b)
    sample_omml3 = '''
    <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <m:rad>
            <m:e>
                <m:r>
                    <m:t>a+b</m:t>
                </m:r>
            </m:e>
        </m:rad>
    </m:oMath>
    '''
    
    latex_result3 = converter.convert_omml_to_latex(sample_omml3)
    print("\n" + "="*50)
    print("OMML Input (simplified view):")
    print("Square root: √(a+b)")
    print("\nLaTeX Output:")
    print(latex_result3)


if __name__ == "__main__":
    main()