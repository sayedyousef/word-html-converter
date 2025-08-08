# ============= word_equation_replacer_fixed.py =============
"""Actually REPLACE equations with plain LaTeX text in Word"""

import win32com.client
from pathlib import Path
import pythoncom
import json
import re

class WordEquationReplacer:
    """Replace Word equations with plain LaTeX text"""
    
    def __init__(self):
        pythoncom.CoInitialize()
        self.word = None
        self.doc = None
        
    def process_document(self, docx_path):
        """Process document - REPLACE equations with LaTeX text"""
        
        docx_path = Path(docx_path).absolute()
        output_path = docx_path.parent / f"{docx_path.stem}_latex_text.docx"
        json_path = docx_path.parent / f"{docx_path.stem}_equations.json"
        
        print(f"\n📁 Processing: {docx_path.name}")
        
        try:
            # Start Word
            print("Starting Word...")
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = False  # Set to True to see what's happening
            
            # Open document
            print("Opening document...")
            self.doc = self.word.Documents.Open(str(docx_path))
            
            # Replace all equations
            equations_data = self._replace_all_equations()
            
            # Save modified document
            print("Saving document with LaTeX text...")
            self.doc.SaveAs2(str(output_path))
            
            # Save JSON
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(equations_data, f, indent=2, ensure_ascii=False)
            
            print(f"\n✅ SUCCESS!")
            print(f"   📄 Word with LaTeX text: {output_path}")
            print(f"   📋 Equations JSON: {json_path}")
            print(f"   ✓ Replaced {len(equations_data)} equations with LaTeX text")
            
            return output_path
            
        finally:
            if self.doc:
                self.doc.Close()
            if self.word:
                self.word.Quit()
            pythoncom.CoUninitialize()
    
    def _replace_all_equations(self):
        """REPLACE each equation with its LaTeX text"""
        
        equations_data = []
        omaths = self.doc.OMaths
        total = omaths.Count
        
        print(f"Found {total} equations to replace")
        
        # IMPORTANT: Process backwards to maintain indices
        for i in range(total, 0, -1):
            try:
                omath = omaths.Item(i)
                
                # Extract LaTeX text BEFORE modifying
                latex_text = self._extract_latex(omath)
                
                # Get the range where equation is
                eq_range = omath.Range
                
                # DELETE the equation object and REPLACE with text
                eq_range.Text = f" {latex_text} "  # Add spaces for clarity
                
                # Now format the replaced text
                eq_range.Font.Name = "Courier New"
                eq_range.Font.Bold = False
                eq_range.Font.Color = 0x0000FF  # Blue color to show it's LaTeX
                eq_range.Shading.BackgroundPatternColor = 0xFFFF00  # Yellow highlight
                
                # Add bookmark
                bookmark_name = f"eq_{i}"
                try:
                    self.doc.Bookmarks.Add(bookmark_name, eq_range)
                except:
                    pass
                
                equations_data.append({
                    'index': i,
                    'latex': latex_text,
                    'bookmark': bookmark_name,
                    'status': 'replaced'
                })
                
                print(f"  ✓ Replaced equation {i} with: {latex_text[:50]}...")
                
            except Exception as e:
                print(f"  ⚠ Equation {i} failed: {str(e)[:50]}")
                
                # Try alternative replacement method
                try:
                    latex_text = self._force_replace_equation(i)
                    equations_data.append({
                        'index': i,
                        'latex': latex_text,
                        'bookmark': f"eq_{i}",
                        'status': 'force_replaced'
                    })
                    print(f"    → Force replaced with: {latex_text[:30]}...")
                except Exception as e2:
                    print(f"    → Could not replace: {e2}")
        
        return equations_data
    
    def _extract_latex(self, omath):
        """Extract LaTeX representation from equation"""
        
        latex = ""
        
        # Method 1: LinearString (best)
        try:
            latex = omath.LinearString
        except:
            pass
        
        # Method 2: Range text
        if not latex:
            try:
                latex = omath.Range.Text
            except:
                pass
        
        # Method 3: BuildUp
        if not latex:
            try:
                omath.BuildUp()
                latex = omath.Range.Text
            except:
                pass
        
        # Clean and convert
        if latex:
            latex = self._clean_to_latex(latex)
        else:
            latex = "[equation]"
        
        return latex
    
    def _force_replace_equation(self, index):
        """Force replacement for problematic equations"""
        
        omath = self.doc.OMaths.Item(index)
        
        # Get any text we can
        latex_text = "[equation]"
        try:
            latex_text = omath.Range.Text or "[equation]"
        except:
            pass
        
        latex_text = self._clean_to_latex(latex_text)
        
        # Select the equation
        omath.Range.Select()
        selection = self.word.Selection
        
        # Type replacement text directly
        selection.TypeText(f" {latex_text} ")
        
        # Format it
        selection.Font.Name = "Courier New"
        selection.Font.Color = 0xFF0000  # Red for force-replaced
        selection.Shading.BackgroundPatternColor = 0xFFFF00
        
        return latex_text
    
    def _clean_to_latex(self, text):
        """Convert text to clean LaTeX format"""
        
        if not text:
            return ""
        
        # Remove control characters
        text = text.replace('\r', ' ').replace('\n', ' ').replace('\x07', '')
        text = text.replace('\x0b', '').replace('\t', ' ')
        
        # Unicode subscripts to LaTeX
        subs = {'₀':'_0', '₁':'_1', '₂':'_2', '₃':'_3', '₄':'_4',
                '₅':'_5', '₆':'_6', '₇':'_7', '₈':'_8', '₉':'_9',
                'ₓ':'_x', 'ᵢ':'_i', 'ⱼ':'_j', 'ₙ':'_n'}
        for u, l in subs.items():
            text = text.replace(u, l)
        
        # Unicode superscripts to LaTeX
        sups = {'⁰':'^0', '¹':'^1', '²':'^2', '³':'^3', '⁴':'^4',
                '⁵':'^5', '⁶':'^6', '⁷':'^7', '⁸':'^8', '⁹':'^9',
                'ⁿ':'^n', 'ⁱ':'^i'}
        for u, l in sups.items():
            text = text.replace(u, l)
        
        # Math symbols to LaTeX
        symbols = {
            '≠': '\\neq', '≤': '\\leq', '≥': '\\geq',
            '∞': '\\infty', '∑': '\\sum', '∏': '\\prod',
            '∫': '\\int', '√': '\\sqrt', '∂': '\\partial',
            'α': '\\alpha', 'β': '\\beta', 'γ': '\\gamma',
            'δ': '\\delta', 'θ': '\\theta', 'π': '\\pi',
            'σ': '\\sigma', 'μ': '\\mu', 'φ': '\\phi',
            '→': '\\rightarrow', '←': '\\leftarrow',
            '⇒': '\\Rightarrow', '⇔': '\\Leftrightarrow',
            '∈': '\\in', '∉': '\\notin', '⊂': '\\subset',
            '∪': '\\cup', '∩': '\\cap', '∀': '\\forall',
            '∃': '\\exists', '±': '\\pm', '×': '\\times',
            '÷': '\\div', '·': '\\cdot'
        }
        for symbol, latex in symbols.items():
            text = text.replace(symbol, latex)
        
        # Pattern fixes
        text = re.sub(r'([a-zA-Z])([0-9]+)', r'\1_{\2}', text)  # x1 → x_{1}
        
        # Clean spaces
        text = ' '.join(text.split())
        
        return text.strip()

# ============= Test it =============
if __name__ == "__main__":
    # Your test file
    test_file = r"D:\Work 3 (20-Oct-24)\2 Side projects May 25\Encyclopedia\articles\مقالات بعد الاخراج\test\الدالة واحد لواحد (جاهزة للنشر).docx"
    
    replacer = WordEquationReplacer()
    output = replacer.process_document(test_file)
    
    print(f"\n🎉 Check the output file - equations should now be PLAIN TEXT in LaTeX format!")
    print(f"   They will be highlighted in YELLOW with BLUE text")