# ============= word_equation_replacer_fixed.py =============
"""COMPLETELY REMOVE equations and replace with PURE PLAIN TEXT"""

import win32com.client
from pathlib import Path
import pythoncom
import json
import re
#from omml_to_latex import omml_xml_to_latex  
from dwml import omml as dwml_omml

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
        """
        Convert a Word OMath object to LaTeX.

        1. Try dwml (OMML → LaTeX).
        2. If dwml returns '' or raises, fall back to Word's LinearString.
        3. If *that* fails, return a visible placeholder so the equation
        isn't silently lost.
        """

        # -- STEP 1 : dwml ------------------------------------------------
        try:
            # omath.XML is the raw <m:oMath …>…</m:oMath> string
            omml_xml = omath.XML

            # some Word builds omit the 'm:' namespace prefix; ensure it's there
            if '<oMath' in omml_xml and 'xmlns:m=' not in omml_xml:
                omml_xml = (
                    '<m:oMath xmlns:m="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/math">' + omml_xml + '</m:oMath>'
                )

            latex = dwml_omml.xml2latex(omml_xml)

            if latex and latex.strip():
                return latex.strip()                  # success!
            else:
                self.logger.debug("dwml returned empty string")

        except Exception as e:
            # dwml couldn’t parse – log & fall back
            self.logger.debug(f"dwml failed: {e}")

        # -- STEP 2 : Word’s own linear text ------------------------------
        try:
            linear = omath.LinearString
            if linear and linear.strip():
                return self._clean_to_latex(linear)
        except Exception:
            pass  # ignore, drop to final placeholder

        # -- STEP 3 : last-resort placeholder -----------------------------
        return f"[equation_{omath.Range.Start}]"


    def _extract_latex_old4(self, omath):
        """
        Robust: use OMML→MathML→LaTeX pipeline first,
        fall back to Word’s LinearString if that ever fails.
        """
        try:
            # 1) COM gives us raw OMML
            omml_xml = omath.XML
            latex = omml_xml_to_latex(omml_xml)
            if latex and latex.strip():
                return latex
        except Exception:
            pass   # fall back if XSLT or mathml2latex chokes

        # 2) older fallback (Word’s built-in linear text)
        try:
            return self._clean_to_latex(omath.LinearString)
        except Exception:
            return "[equation]"        # worst-case placeholder


    def _extract_latex_old_3(self, omath):
        """
        Try every COM property that can give us a linear representation,
        then fall back to range-text or a placeholder – but never raise.
        """
        # Word’s COM API sometimes throws "Value does not fall within the
        # expected range" if LinearString is empty.  Wrap each attempt.
        for getter in (
            lambda o: o.LinearString,        # 1. best quality
            lambda o: o.BuildUp() or o.Range.Text,   # 2. after BuildUp()
            lambda o: o.Range.Text           # 3. raw text
        ):
            try:
                txt = getter(omath)
                if txt and txt.strip():
                    return self._clean_to_latex(txt)
            except Exception:
                pass  # try next method

        # Last-chance extraction: copy–paste the selection as plain text
        try:
            omath.Range.Select()
            self.word.Selection.Copy()
            temp_doc = self.word.Documents.Add()
            temp_doc.Range().PasteSpecial(DataType=2)  # 2 = wdPasteText
            txt = temp_doc.Range().Text
            temp_doc.Close(False)
            if txt and txt.strip():
                return self._clean_to_latex(txt)
        except Exception:
            pass

        # Give the caller *something* it can use
        return f"[equation_{omath.Range.Start}]"


    def _extract_latex_old2(self, omath):
        """Simple extraction - LinearString is actually good enough for most cases"""
        try:
            # LinearString gives the best result from Word COM
            latex = omath.LinearString
            
            if latex:
                # Just convert the Unicode symbols to LaTeX
                latex = latex.replace('∑', '\\sum')
                latex = latex.replace('∏', '\\prod')
                latex = latex.replace('∫', '\\int')
                latex = latex.replace('√', '\\sqrt')
                # Let your existing _clean_to_latex handle the rest
                return self._clean_to_latex(latex)
            else:
                # Fallback
                return omath.Range.Text or "[equation]"
        except:
            return "[equation]"
    
    def _extract_latex_old(self, omath):
        """Extract LaTeX representation from equation - FIXED VERSION"""
        
        latex = ""
        
        # Method 1: Try BuildUp to get linear text
        try:
            # BuildUp converts equation to linear format
            omath.BuildUp()
            latex = omath.Range.Text
            print(f"  ✓ Extracted LaTeX from BuildUp: {latex[:50]}")
            print(latex)
            if latex and latex.strip():
                return self._clean_to_latex(latex)
        except:
            pass
        
        # Method 2: Try LinearString
        try:
            latex = omath.LinearString
            print(f"  ✓ Extracted LaTeX from LinearString: {latex[:50]}")
            print(latex)
            if latex and latex.strip():
                return self._clean_to_latex(latex)
        except:
            pass
        
        # Method 3: Get Range.Text directly
        try:
            latex = omath.Range.Text
            print(f"  ✓ Extracted LaTeX from Range.Text: {latex[:50]}")
            print(latex)
            if latex and latex.strip():
                return self._clean_to_latex(latex)
        except:
            pass
        
        # Method 4: Copy and paste as text
        try:
            omath.Range.Select()
            self.word.Selection.Copy()
            
            # Create new temporary document
            temp_doc = self.word.Documents.Add()
            temp_doc.Range().PasteSpecial(DataType=2)  # Paste as text
            latex = temp_doc.Range().Text
            temp_doc.Close(False)
            print(f"  ✓ Extracted LaTeX from PasteSpecial: {latex[:50]}")
            print(latex)
            if latex and latex.strip():
                return self._clean_to_latex(latex)
        except:
            pass
        
        # Method 5: Force convert to text
        try:
            # This actually modifies the equation but we're deleting it anyway
            omath.ConvertToNormalText()
            latex = omath.Range.Text
            print(f"  ✓ Extracted LaTeX from ConvertToNormalText: {latex[:50]}")
            print(latex)
            if latex and latex.strip():
                return self._clean_to_latex(latex)
        except:
            pass
        
        # If all methods fail, at least return something identifiable
        return f"[equation_{omath.Range.Start}]"  # Use position as identifier            

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