"""
Complete OMML to LaTeX Parser - All Issues Fixed
Properly handles matrices, integrals, binomials, symbols with spacing
"""

import os
import zipfile
import re
from lxml import etree

# Symbol mapping
MATH_SYMBOLS = {
    '≠': r'\neq', '≤': r'\leq', '≥': r'\geq', '±': r'\pm', '×': r'\times',
    '÷': r'\div', '·': r'\cdot', '≈': r'\approx', '≡': r'\equiv', '∼': r'\sim',
    '∈': r'\in', '∉': r'\notin', '⊂': r'\subset', '⊆': r'\subseteq',
    '∪': r'\cup', '∩': r'\cap', '∅': r'\emptyset', '∧': r'\land', '∨': r'\lor',
    '¬': r'\neg', '∀': r'\forall', '∃': r'\exists', '→': r'\rightarrow',
    '←': r'\leftarrow', '↔': r'\leftrightarrow', '⇒': r'\Rightarrow',
    'α': r'\alpha', 'β': r'\beta', 'γ': r'\gamma', 'δ': r'\delta', 'ε': r'\epsilon',
    'θ': r'\theta', 'λ': r'\lambda', 'μ': r'\mu', 'π': r'\pi', 'σ': r'\sigma',
    'τ': r'\tau', 'φ': r'\phi', 'ψ': r'\psi', 'ω': r'\omega', 
    'Γ': r'\Gamma', 'Δ': r'\Delta', 'Σ': r'\Sigma', 'Ω': r'\Omega',
    '∂': r'\partial', '∇': r'\nabla', '∑': r'\sum', '∏': r'\prod', '∫': r'\int',
    '∞': r'\infty', '√': r'\sqrt', '∠': r'\angle', '⊥': r'\perp', '∥': r'\parallel',
    '…': r'\ldots', '∴': r'\therefore', '∵': r'\because', '°': r'^\circ',
    'υ': r'\upsilon', 'ϒ': r'\Upsilon',
    'ⅆ': r'\, d',  # Differential d with thin space before it
    '∓': r'\mp',  # Missing minus-plus
}

FUNCTION_NAMES = {
    'sin': r'\sin', 'cos': r'\cos', 'tan': r'\tan', 'sec': r'\sec',
    'csc': r'\csc', 'cot': r'\cot', 'arcsin': r'\arcsin', 'arccos': r'\arccos',
    'sinh': r'\sinh', 'cosh': r'\cosh', 'tanh': r'\tanh', 'log': r'\log',
    'ln': r'\ln', 'exp': r'\exp', 'lim': r'\lim', 'sup': r'\sup', 'inf': r'\inf',
    'min': r'\min', 'max': r'\max', 'det': r'\det', 'dim': r'\dim',
}

class DirectOmmlToLatex:
    def __init__(self):
        self.ns = {
            'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        }
            
    def smart_symbol_convert(self, text):
        """Convert symbols with smart spacing"""
        result = []
        i = 0
        while i < len(text):
            found = False
            for symbol, latex in MATH_SYMBOLS.items():
                if text[i:i+len(symbol)] == symbol:
                    result.append(latex)
                    # General rule: LaTeX commands need space before letters
                    if i + len(symbol) < len(text):
                        next_char = text[i + len(symbol)]
                        # If it's a LaTeX command and next is letter, add space
                        if latex.startswith('\\') and next_char.isalpha():
                            result.append(' ')
                    i += len(symbol)
                    found = True
                    break
            if not found:
                result.append(text[i])
                i += 1
        return ''.join(result)


    
    def convert_function_names(self, text):
        """Convert function names to LaTeX"""
        if text.startswith('\\'):
            return text
        for func, latex_func in FUNCTION_NAMES.items():
            text = re.sub(r'\b' + re.escape(func) + r'(?=\s|\(|$)', 
                         lambda m: latex_func, text)
        return text
    
    def clean_output(self, latex):
        """Clean LaTeX output carefully"""
        # Skip cleaning for certain patterns
        if any(cmd in latex for cmd in ['\\binom', '\\left', '\\right', '\\begin']):
            # Only do minimal cleaning for complex structures
            latex = re.sub(r'\s+_', '_', latex)
            latex = re.sub(r'\s+\^', '^', latex)
            # Fix partial derivatives
            latex = re.sub(r'(\\partial)([a-zA-Z])', r'\1 \2', latex)
            # Fix missing braces in fractions
            latex = re.sub(r'\\frac([a-zA-Z0-9])\{', r'\\frac{\1}{', latex)
            return latex
            
        # Regular cleaning for simple content
        # Don't remove braces from single characters after backslash commands
        latex = re.sub(r'(?<!\\[a-zA-Z])\{([a-zA-Z0-9])\}', r'\1', latex)
        latex = re.sub(r'\{\{([^}]+)\}\}', r'{\1}', latex)
        latex = re.sub(r'\s+_', '_', latex)
        latex = re.sub(r'\s+\^', '^', latex)
        # Fix partial derivatives
        latex = re.sub(r'(\\partial)([a-zA-Z])', r'\1 \2', latex)
        # Fix missing braces in fractions
        latex = re.sub(r'\\frac([a-zA-Z0-9])\{', r'\\frac{\1}{', latex)
        return latex
    
    def parse(self, elem):
        """Main parsing function"""
        if elem is None:
            return ''
        
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
        handler = getattr(self, f'parse_{tag}', self.parse_default)
        return handler(elem)
    
    def parse_oMath(self, elem):
        latex = ''.join(self.parse(child) for child in elem)
        latex = self.clean_output(latex)
        latex = self.apply_post_processing(latex)
        return latex
    
    def parse_oMathPara(self, elem):
        return ''.join(self.parse(child) for child in elem)

    def parse_r(self, elem):
        """Run element with smart symbol handling"""
        texts = elem.xpath('.//m:t/text()', namespaces=self.ns)
        text = ''.join(texts)
        
        # Handle minus sign first
        text = text.replace('−', '-')
        
        # FIX: Handle differential d (ⅆ) with proper LaTeX spacing
        # Pattern 'rⅆrⅆ' should become 'r \, dr \, d'
        text = re.sub(r'([a-z])ⅆ([a-z])ⅆ', r'\1 \, d\2 \, d', text)

        
        # Handle single differential like 'xⅆ' -> 'x \, d'
        text = re.sub(r'([a-z])ⅆ', r'\1 \, d', text)
        
        # Also handle regular 'd' as differential when it follows a variable
        # This catches cases where 'd' is already regular 'd' not 'ⅆ'
        text = re.sub(r'([a-z])d([a-z])d\b', r'\1 \, d\2 \, d', text)
        text = re.sub(r'([a-z])d([αβγδεζηθικλμνξοπρστυφχψω])', r'\1 \, d\2', text)
        
        # Convert symbols with smart spacing
        text = self.smart_symbol_convert(text)
        
        # FIX: Add space after Greek letters when followed by variables
        # This fixes γz → \gamma z in superscripts
        text = re.sub(r'(\\gamma|\\alpha|\\beta|\\delta|\\theta|\\sigma)([a-z])', r'\1 \2', text)

        # Convert function names
        text = self.convert_function_names(text)
        
        return text


    def parse_f(self, elem):
        """Fraction with proper binomial detection"""
        num_elem = elem.find('.//m:num', self.ns)
        den_elem = elem.find('.//m:den', self.ns)
        
        num = self.parse(num_elem) if num_elem is not None else ''
        den = self.parse(den_elem) if den_elem is not None else ''
        
        # Strip spaces but preserve content
        num = num.strip()
        den = den.strip()
        
        # Special case for 1/2 type fractions in superscripts
        if num in ['1', '2', '3'] and den in ['2', '3', '4']:
            return f'\\frac{{{num}}}{{{den}}}'
        
        # Binomial coefficient detection - only for n and k pattern (common binomial notation)
        if (len(num) == 1 and len(den) == 1 and
            #num == 'n' and den == 'k'):
            #return f'\\binom{{{num}}}{{{den}}}'
            #if (len(num) == 1 and len(den) == 1 and
            num.isalpha() and den.isalpha() and
            ((num == 'n' and den == 'k') or 
            (elem.getparent() is not None and 
            elem.getparent().tag.endswith('d')))):  # Check if inside delimiters
            return f'\\binom{{{num}}}{{{den}}}'
        
        # Regular fraction - ensure braces are always present
        return f'\\frac{{{num}}}{{{den}}}'
    
    def parse_sSup(self, elem):
        """Superscript - handle complex nested structures"""
        base_elem = elem.find('m:e', self.ns)
        sup_elem = elem.find('m:sup', self.ns)
        
        base = self.parse(base_elem) if base_elem is not None else ''
        sup = self.parse(sup_elem) if sup_elem is not None else ''
        
        # For complex bracketed expressions, check for duplicate content
        if base.startswith('\\left['):
            # Check if the base already contains nested integrals
            # Count occurrences of key elements
            integral_count = base.count('\\int')
            if integral_count > 2:  # More than expected means duplication
                # Try to extract just the first integral expression
                parts = base.split('\\int')
                if len(parts) > 3:
                    # Reconstruct with just the needed parts
                    base = '\\left[' + '\\int'.join(parts[:3]) + '\\right]'
        
        # Clean the base for simple cases
        if not any(cmd in base for cmd in ['\\binom', '\\left', '\\right', '\\begin']):
            base = self.clean_output(base)
        
        return f'{base}^{{{sup}}}'
    
    def parse_sSub(self, elem):
        """Subscript"""
        base_elem = elem.find('m:e', self.ns)
        sub_elem = elem.find('m:sub', self.ns)
        
        base = self.parse(base_elem) if base_elem is not None else ''
        sub = self.parse(sub_elem) if sub_elem is not None else ''
        
        base = self.clean_output(base)
        return f'{base}_{{{sub}}}'
    
    def parse_sSubSup(self, elem):
        """Sub and superscript"""
        base_elem = elem.find('m:e', self.ns)
        sub_elem = elem.find('m:sub', self.ns)
        sup_elem = elem.find('m:sup', self.ns)
        
        base = self.parse(base_elem) if base_elem is not None else ''
        sub = self.parse(sub_elem) if sub_elem is not None else ''
        sup = self.parse(sup_elem) if sup_elem is not None else ''
        
        base = self.clean_output(base)
        return f'{base}_{{{sub}}}^{{{sup}}}'
    
    def parse_nary(self, elem):
        """N-ary operations"""
        chr_elem = elem.find('.//m:naryPr/m:chr', self.ns)
        
        if chr_elem is not None:
            op_val = chr_elem.get(f'{{{self.ns["m"]}}}val', '∫')
        else:
            op_val = '∫'
        
        operator = self.smart_symbol_convert(op_val)
        
        sub_elem = elem.find('m:sub', self.ns)
        sup_elem = elem.find('m:sup', self.ns)
        expr_elem = elem.find('m:e', self.ns)
        
        result = operator
        if sub_elem is not None:
            result += f'_{{{self.parse(sub_elem)}}}'
        if sup_elem is not None:
            result += f'^{{{self.parse(sup_elem)}}}'
        if expr_elem is not None:
            result += f' {self.parse(expr_elem)}'
        
        return result
    
    def parse_rad(self, elem):
        """Radical"""
        deg_elem = elem.find('m:deg', self.ns)
        expr_elem = elem.find('m:e', self.ns)
        
        expr = self.parse(expr_elem) if expr_elem is not None else ''
        
        # Check if degree is hidden
        deg_hide = elem.find('.//m:degHide', self.ns)
        if deg_hide is not None and deg_hide.get(f'{{{self.ns["m"]}}}val') == '1':
            return f'\\sqrt{{{expr}}}'
        
        if deg_elem is not None:
            deg_text = self.parse(deg_elem)
            if deg_text and deg_text.strip():
                return f'\\sqrt[{deg_text}]{{{expr}}}'
        
        return f'\\sqrt{{{expr}}}'
    
    def parse_d(self, elem):
        """Delimiters - handle all types properly"""
        beg_chr = elem.find('.//m:begChr', self.ns)
        end_chr = elem.find('.//m:endChr', self.ns)
        
        open_d = '('
        close_d = ')'
        
        if beg_chr is not None:
            open_d = beg_chr.get(f'{{{self.ns["m"]}}}val', '(')
        if end_chr is not None:
            close_d = end_chr.get(f'{{{self.ns["m"]}}}val', ')')
        
        # Get direct e children
        e_children = []
        for child in elem:
            if child.tag.endswith('e'):
                e_children.append(child)
        
        if not e_children:
            return ''
        
        # Check first e child for special structures
        first_e = e_children[0]
        for grandchild in first_e:
            if grandchild.tag.endswith('m'):
                # Matrix
                matrix_type = 'pmatrix'
                if open_d == '[':
                    matrix_type = 'bmatrix'
                elif open_d == '{':
                    matrix_type = 'Bmatrix'
                elif open_d == '|':
                    matrix_type = 'vmatrix'
                return self.parse_matrix(grandchild, matrix_type)
            elif grandchild.tag.endswith('eqArr'):
                # Piecewise
                content = self.parse(grandchild)
                if open_d == '{' and (not close_d or close_d == ''):
                    return f'\\begin{{cases}} {content} \\end{{cases}}'
                return content
        
        # Parse the content
        inner = self.parse(first_e)
        
        # Apply delimiters based on type
        if open_d == '(' and close_d == ')':
            return f'\\left({inner}\\right)'
        elif open_d == '[' and close_d == ']':
            return f'\\left[{inner}\\right]'
        elif open_d == '{' and close_d == '}':
            return f'\\left\\{{{inner}\\right\\}}'
        elif open_d == '|' and close_d == '|':
            return f'\\left|{inner}\\right|'
        else:
            return f'{open_d}{inner}{close_d}'
    
    def parse_matrix(self, elem, matrix_type='pmatrix'):
        """Parse matrix elements correctly"""
        rows = []
        
        # Process each row (mr element)
        for child in elem:
            if child.tag.endswith('mr'):
                cols = []
                # Process each cell (e element) in the row
                for cell in child:
                    if cell.tag.endswith('e'):
                        cell_content = self.parse(cell)
                        if cell_content:
                            cols.append(cell_content)
                
                # Only add non-empty rows
                if cols:
                    rows.append(' & '.join(cols))
        
        # Join rows with line breaks
        if rows:
            content = ' \\\\ '.join(rows)
            return f'\\begin{{{matrix_type}}} {content} \\end{{{matrix_type}}}'
        
        return ''
    
    def parse_m(self, elem):
        """Matrix without delimiters"""
        return self.parse_matrix(elem, 'matrix')
    
    def parse_func(self, elem):
        """Functions with proper handling"""
        fname_elem = elem.find('m:fName', self.ns)
        arg_elem = elem.find('m:e', self.ns)
        
        # Handle limit with subscript
        if fname_elem is not None:
            limlower = fname_elem.find('.//m:limLow', self.ns)
            if limlower is not None:
                fname_parsed = self.parse(fname_elem)
                arg_parsed = self.parse(arg_elem) if arg_elem is not None else ''
                
                # If argument is just the expression after limit
                if arg_parsed and not ('\\lim' in arg_parsed):
                    return f'{fname_parsed} {arg_parsed}'
                else:
                    # Argument might be empty or contain limit already
                    return fname_parsed
        
        # Regular functions
        fname = self.parse(fname_elem) if fname_elem is not None else ''
        arg = self.parse(arg_elem) if arg_elem is not None else ''
        
        # Convert function names
        if fname and not fname.startswith('\\'):
            fname = self.convert_function_names(fname)
        
        # Limits don't get parentheses
        if fname and 'lim' in fname.lower():
            if arg:
                return f'{fname} {arg}'
            return fname
        
        # Other functions
        if fname and arg:
            return f'{fname}({arg})'
        elif fname:
            return fname
        else:
            return arg or ''
    
    def parse_limLow(self, elem):
        """Limit lower - for limits with subscripts"""
        base_elem = elem.find('m:e', self.ns)
        lim_elem = elem.find('m:lim', self.ns)
        
        base = self.parse(base_elem) if base_elem is not None else ''
        lim = self.parse(lim_elem) if lim_elem is not None else ''
        
        # Convert lim to LaTeX if needed
        if base == 'lim':
            base = '\\lim'
        elif not base.startswith('\\'):
            base = self.convert_function_names(base)
        
        return f'{base}_{{{lim}}}'
    
    def parse_acc(self, elem):
        """Accents"""
        chr_elem = elem.find('.//m:accPr/m:chr', self.ns)
        base_elem = elem.find('m:e', self.ns)
        
        base = self.parse(base_elem) if base_elem is not None else ''
        
        if chr_elem is not None:
            acc_val = chr_elem.get(f'{{{self.ns["m"]}}}val', '')
            accent_map = {
                '̂': 'hat', '̃': 'tilde', '̄': 'bar',
                '̇': 'dot', '̈': 'ddot', '⃗': 'vec',
            }
            latex_acc = accent_map.get(acc_val, 'hat')
            return f'\\{latex_acc}{{{base}}}'
        
        return f'\\hat{{{base}}}'
    
    def parse_eqArr(self, elem):
        """Equation array for piecewise functions"""
        parts = []
        for child in elem:
            if child.tag.endswith('e'):
                part = self.parse(child)
                if part and part.strip():
                    parts.append(part.strip())
        
        # Format for cases environment
        formatted_parts = []
        for part in parts:
            # Look for patterns like "a, n odd" or "a, &n even"
            if ',' in part:
                pieces = part.split(',', 1)  # Split on first comma only
                value = pieces[0].strip()
                condition = pieces[1].strip() if len(pieces) > 1 else ''
                
                # Remove leading & if present
                if condition.startswith('&'):
                    condition = condition[1:].strip()
                
                # Format as value & condition
                if condition:
                    if 'odd' in condition or 'even' in condition:
                        formatted_parts.append(f'{value}, & \\text{{{condition}}}')
                    else:
                        formatted_parts.append(f'{value}, & {condition}')
                else:
                    formatted_parts.append(value)
            else:
                formatted_parts.append(part)
        
        return ' \\\\ '.join(formatted_parts)
    
    def parse_default(self, elem):
        """Default handler - process children sequentially"""
        results = []
        for child in elem:
            result = self.parse(child)
            if result:
                results.append(result)
        return ''.join(results)
    
    def parse_e(self, elem):
        """Element container"""
        results = []
        for child in elem:
            result = self.parse(child)
            if result:
                results.append(result)
        return ''.join(results)
    # Aliases for simple pass-through elements
    parse_num = parse_default
    parse_den = parse_default
    parse_sub = parse_default
    parse_sup = parse_default
    parse_lim = parse_default
    parse_limUpp = parse_limLow
    parse_mr = parse_default


    def apply_post_processing(self, latex):
        """Apply all post-processing fixes"""
        # All the fixes from process_word_document
        latex = re.sub(r'\\binom([a-zA-Z])([a-zA-Z])', r'\\binom{\1}{\2}', latex)
        latex = re.sub(r'(e\^{[^}]+}[a-z]+)(.*?)\1', r'\1\2', latex)
        latex = re.sub(r'([a-zA-Z]+)\\left\(([^)]+)\\right\)\1', r'\1\\left(\2\\right)', latex)
        latex = re.sub(r'\\partial([a-zA-Z])', r'\\partial \1', latex)
        latex = re.sub(r'\\upsilon([a-zA-Z])', r'\\upsilon \1', latex)
        latex = re.sub(r'\\gamma([a-zA-Z])', r'\\gamma \1', latex)
        latex = re.sub(r'\\rightarrow([A-Z][a-z])', r'\\rightarrow \1', latex)
        latex = latex.replace('⋅', r'\cdot')
        latex = re.sub(r'(\\lim[^}]*})\s*\\lim\s', r'\1 ', latex)
        latex = re.sub(r'(\\exists|\\forall)([a-zA-Z])', r'\1 \2', latex)
        latex = re.sub(r'\\left\(\\binom\{([^}]+)\}\{([^}]+)\}\\right\)', r'\\binom{\1}{\2}', latex)
        latex = re.sub(r'\\cdot([A-Za-z])', r'\\cdot \1', latex)
        latex = re.sub(r'(\\approx|\\equiv|\\sim)(\d)', r'\1 \2', latex)
        latex = re.sub(r'\\cdot([A-Za-z])', r'\\cdot \1', latex)
        return latex
