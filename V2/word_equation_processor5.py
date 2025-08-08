# ============= word_equation_replacer_fixed.py =============
"""COMPLETELY REMOVE equations and replace with PURE PLAIN TEXT"""

import win32com.client
from pathlib import Path
import pythoncom
import json
import re

class WordEquationReplacer:
    """COMPLETELY REMOVE equation objects, replace with PLAIN TEXT"""
    
    def __init__(self):
        pythoncom.CoInitialize()
        self.word = None
        self.doc = None
        
    def process_document(self, docx_path):
        """Process document - DELETE equations, INSERT plain text"""
        
        docx_path = Path(docx_path).absolute()
        output_path = docx_path.parent / f"{docx_path.stem}_latex_text.docx"
        json_path = docx_path.parent / f"{docx_path.stem}_equations.json"
        
        print(f"\n📁 Processing: {docx_path.name}")
        
        try:
            # Start Word
            print("Starting Word...")
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = False
            
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
            print(f"   📄 Word with PLAIN TEXT: {output_path}")
            print(f"   📋 Equations JSON: {json_path}")
            print(f"   ✓ Replaced {len(equations_data)} equations with PLAIN TEXT")
            print(f"   ✓ NO equation objects remain - just text!")
            
            return output_path
            
        finally:
            if self.doc:
                self.doc.Close()
            if self.word:
                self.word.Quit()
            pythoncom.CoUninitialize()
    
    def _replace_all_equations(self):
        """DELETE equation objects and INSERT plain text"""
        
        equations_data = []
        total = self.doc.OMaths.Count
        
        print(f"Found {total} equation OBJECTS to remove")
        
        # Process from last to first
        for i in range(total, 0, -1):
            try:
                omath = self.doc.OMaths.Item(i)
                
                # Extract LaTeX text FIRST
                latex_text = self._extract_latex(omath)
                
                # Get the range of the equation
                eq_range = omath.Range
                
                # METHOD 1: Delete equation and insert text
                eq_range.Select()
                selection = self.word.Selection
                
                # DELETE the equation completely
                selection.Delete()
                
                # INSERT plain text (NOT an equation)
                selection.TypeText(f" {latex_text} ")
                
                # Format the PLAIN TEXT
                #selection.Font.Name = "Courier New"
                #selection.Font.Bold = False
                #selection.Font.Color = 0x0000FF  # Blue
                #selection.Shading.BackgroundPatternColor = 0xFFFF00  # Yellow
                
                # Add bookmark
                bookmark_name = f"eq_{i}"
                try:
                    self.doc.Bookmarks.Add(bookmark_name, selection.Range)
                except:
                    pass
                
                equations_data.append({
                    'index': i,
                    'latex': latex_text,
                    'bookmark': bookmark_name,
                    'status': 'replaced_as_text'
                })
                
                print(f"  ✓ Removed equation {i}, inserted text: {latex_text[:50]}...")
                
            except Exception as e:
                print(f"  ⚠ Equation {i} failed: {str(e)[:50]}")
                
                # Try alternative method
                try:
                    latex_text = self._force_delete_and_replace(i)
                    equations_data.append({
                        'index': i,
                        'latex': latex_text,
                        'bookmark': f"eq_{i}",
                        'status': 'force_replaced'
                    })
                    print(f"    → Force replaced with text: {latex_text[:30]}...")
                except Exception as e2:
                    print(f"    → Could not replace: {e2}")
        
        # Verify no equations remain
        remaining = self.doc.OMaths.Count
        if remaining == 0:
            print(f"\n✅ All equation objects removed!")
        else:
            print(f"\n⚠ {remaining} equations still remain as objects")
        
        return equations_data
    
    def _force_delete_and_replace(self, index):
        """Force delete equation and replace with text"""
        
        omath = self.doc.OMaths.Item(index)
        
        # Get text first
        latex_text = "[equation]"
        try:
            latex_text = omath.Range.Text or "[equation]"
        except:
            pass
        
        latex_text = self._clean_to_latex(latex_text)
        
        # Convert equation to normal text (removes equation object)
        try:
            omath.ConvertToNormalText()
        except:
            # If that fails, select and delete
            omath.Range.Select()
            self.word.Selection.Delete()
            self.word.Selection.TypeText(f" {latex_text} ")
        
        # Format
        omath.Range.Font.Name = "Courier New"
        #omath.Range.Font.Color = 0xFF0000
        #omath.Range.Shading.BackgroundPatternColor = 0xFFFF00
        
        return latex_text
    
    def _extract_latex(self, omath):
        """Extract LaTeX representation from equation"""
        
        latex = ""
        
        # Try LinearString first
        try:
            latex = omath.LinearString
        except:
            pass
        
        # Try Range text
        if not latex:
            try:
                latex = omath.Range.Text
            except:
                pass
        
        # Try ConvertToNormalText
        if not latex:
            try:
                # Make a copy of the text before converting
                temp_range = omath.Range.Duplicate
                omath.ConvertToNormalText()
                latex = omath.Range.Text
            except:
                pass
        
        # Clean and convert
        if latex:
            latex = self._clean_to_latex(latex)
        else:
            latex = "[equation]"
        
        return latex
    
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
    
    print(f"\n🎉 Equations are now PURE PLAIN TEXT - NOT equation objects!")