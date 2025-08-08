# ============= word_com_equation_replacer.py =============
"""Use Windows Word COM to replace equations - preserves document perfectly"""
import win32com.client
from pathlib import Path
import pythoncom
from logger import setup_logger

logger = setup_logger("word_com_replacer")

class WordCOMEquationReplacer:
    """Use Word itself to replace equations - perfect preservation"""
    
    def __init__(self):
        # Initialize COM
        pythoncom.CoInitialize()
        self.word = None
        self.doc = None
        
    def process_document(self, docx_path, output_path=None):
        """Open document in Word and replace equations"""
        
        docx_path = Path(docx_path).absolute()
        
        # Make sure output is in same folder as input
        if not output_path:
            output_path = docx_path.parent / f"{docx_path.stem}_equations_text.docx"
        else:
            output_path = Path(output_path).absolute()
        
        print(f"\n📁 Input file: {docx_path}")
        print(f"📁 Output will be saved to: {output_path}")
        print(f"📁 JSON will be saved to: {output_path.parent / f'{docx_path.stem}_equations.json'}\n")
        
        try:
            # Start Word
            logger.info("Starting Word application...")
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = False
            
            # Open document
            logger.info(f"Opening {docx_path.name}...")
            self.doc = self.word.Documents.Open(str(docx_path))
            
            # Process equations
            equation_count = self._replace_equations()
            
            # Save as new document with FULL PATH
            logger.info(f"Saving to {output_path}...")
            self.doc.SaveAs2(str(output_path))
            
            # Also save JSON with equations data
            json_path = output_path.parent / f"{docx_path.stem}_equations.json"
            
            print(f"\n✅ SUCCESS! Files saved:")
            print(f"   📄 Word doc: {output_path}")
            print(f"   📋 JSON data: {json_path}")
            
            # Check if files exist
            if output_path.exists():
                print(f"   ✓ Word file size: {output_path.stat().st_size:,} bytes")
            else:
                print(f"   ❌ Word file NOT FOUND!")
                
            if json_path.exists():
                print(f"   ✓ JSON file size: {json_path.stat().st_size:,} bytes")
            else:
                print(f"   ❌ JSON file NOT FOUND!")
            
            return output_path
            
        except Exception as e:
            logger.error(f"Error: {e}")
            raise
            
        finally:
            self._cleanup()
    
# ============= Fixed version that handles subscripts properly =============
class WordCOMEquationReplacer:
    """Fixed version with better equation handling"""
    
    def _replace_equations(self):
        """Find and replace all equations - handles multi-line properly"""
        
        equations_replaced = 0
        equations_data = []
        
        # Get all OMath objects
        omaths = self.doc.OMaths
        logger.info(f"Found {omaths.Count} equations")
        
        for i in range(omaths.Count, 0, -1):
            try:
                omath = omaths.Item(i)
                
                # Method 1: Try to get LinearString (best for LaTeX)
                latex_text = None
                try:
                    # LinearString gives the linear format
                    latex_text = omath.LinearString
                    logger.debug(f"Got LinearString: {latex_text}")
                except:
                    pass
                
                # Method 2: Try BuildUp text
                if not latex_text:
                    try:
                        # First convert to BuildUp format
                        omath.BuildUp()
                        latex_text = omath.Range.Text
                        logger.debug(f"Got BuildUp text: {latex_text}")
                    except:
                        pass
                
                # Method 3: Convert to normal text
                if not latex_text:
                    try:
                        # Make a copy of the range
                        temp_range = omath.Range.Duplicate
                        
                        # Convert equation to normal text
                        omath.ConvertToNormalText()
                        
                        # Get the text
                        latex_text = omath.Range.Text
                        logger.debug(f"Got normal text: {latex_text}")
                    except:
                        # Get raw text
                        latex_text = omath.Range.Text
                
                # Clean up the text
                if latex_text:
                    # Remove control characters
                    latex_text = latex_text.replace('\r', '')
                    latex_text = latex_text.replace('\n', '')
                    latex_text = latex_text.replace('\x07', '')  # Bell character
                    latex_text = latex_text.replace('\x0b', '')  # Vertical tab
                    
                    # Fix subscripts (x₁ becomes x_1)
                    latex_text = self._fix_subscripts(latex_text)
                    
                    # Fix superscripts (x² becomes x^2)
                    latex_text = self._fix_superscripts(latex_text)
                    
                    # Convert symbols
                    latex_text = self._convert_symbols(latex_text)
                    
                    # Clean whitespace
                    latex_text = ' '.join(latex_text.split())
                
                if not latex_text or latex_text.strip() == '':
                    latex_text = "[equation]"
                
                # Create bookmark name
                bookmark_name = f"eq_{i}"
                
                # The equation is now converted to text, just format it
                equation_range = omath.Range
                equation_range.Font.Name = "Courier New"
                equation_range.Font.Color = 0x333333
                equation_range.Shading.BackgroundPatternColor = 0xF0F0F0
                
                # Add bookmark
                try:
                    self.doc.Bookmarks.Add(bookmark_name, equation_range)
                except:
                    pass
                
                equations_replaced += 1
                equations_data.append({
                    'index': i,
                    'latex': latex_text,
                    'bookmark': bookmark_name
                })
                
                logger.debug(f"Replaced equation {i}: {latex_text}")
                
            except Exception as e:
                logger.warning(f"Could not process equation {i}: {e}")
                
                # Try alternative method for problematic equations
                try:
                    self._handle_problematic_equation(i, equations_data)
                except:
                    pass
        
        # Save equation data
        self._save_equation_data(equations_data)
        
        return equations_replaced
    
    def _handle_problematic_equation(self, index, equations_data):
        """Handle equations with paragraph marks"""
        
        omath = self.doc.OMaths.Item(index)
        
        # Select the equation
        omath.Range.Select()
        selection = self.word.Selection
        
        # Copy to clipboard
        selection.Copy()
        
        # Delete the equation
        selection.Delete()
        
        # Paste as text
        selection.PasteSpecial(DataType=2)  # wdPasteText
        
        # Get the pasted text
        text = selection.Text
        
        # Clean it up
        text = text.replace('\r', '').replace('\n', ' ')
        text = self._fix_subscripts(text)
        text = self._fix_superscripts(text)
        text = self._convert_symbols(text)
        
        # Format it
        selection.Font.Name = "Courier New"
        selection.Font.Color = 0x333333
        selection.Shading.BackgroundPatternColor = 0xF0F0F0
        
        equations_data.append({
            'index': index,
            'latex': text,
            'bookmark': f"eq_{index}"
        })
        
        logger.info(f"Handled problematic equation {index}")
    
    def _fix_subscripts(self, text):
        """Convert Unicode subscripts to LaTeX format"""
        
        # Unicode subscript mapping
        subscripts = {
            '₀': '_0', '₁': '_1', '₂': '_2', '₃': '_3', '₄': '_4',
            '₅': '_5', '₆': '_6', '₇': '_7', '₈': '_8', '₉': '_9',
            'ₐ': '_a', 'ₑ': '_e', 'ₒ': '_o', 'ₓ': '_x', 'ₕ': '_h',
            'ₖ': '_k', 'ₗ': '_l', 'ₘ': '_m', 'ₙ': '_n', 'ₚ': '_p',
            'ₛ': '_s', 'ₜ': '_t'
        }
        
        for unicode_char, latex in subscripts.items():
            text = text.replace(unicode_char, latex)
        
        # Fix patterns like x1 that should be x_1
        import re
        text = re.sub(r'([a-zA-Z])([0-9]+)', r'\1_{\2}', text)
        
        return text
    
    def _fix_superscripts(self, text):
        """Convert Unicode superscripts to LaTeX format"""
        
        # Unicode superscript mapping
        superscripts = {
            '⁰': '^0', '¹': '^1', '²': '^2', '³': '^3', '⁴': '^4',
            '⁵': '^5', '⁶': '^6', '⁷': '^7', '⁸': '^8', '⁹': '^9',
            'ⁿ': '^n', 'ⁱ': '^i'
        }
        
        for unicode_char, latex in superscripts.items():
            text = text.replace(unicode_char, latex)
        
        return text
    
    def _convert_symbols(self, text):
        """Convert mathematical symbols to LaTeX"""
        
        symbols = {
            '≠': '\\neq',
            '≤': '\\leq',
            '≥': '\\geq',
            '∞': '\\infty',
            '∑': '\\sum',
            '∏': '\\prod',
            '∫': '\\int',
            '√': '\\sqrt',
            'α': '\\alpha',
            'β': '\\beta',
            'γ': '\\gamma',
            'δ': '\\delta',
            'π': '\\pi',
            'σ': '\\sigma',
            'θ': '\\theta',
            '∈': '\\in',
            '∉': '\\notin',
            '⊂': '\\subset',
            '⊆': '\\subseteq',
            '∪': '\\cup',
            '∩': '\\cap',
            '∀': '\\forall',
            '∃': '\\exists',
            '→': '\\rightarrow',
            '⇒': '\\Rightarrow',
            '↔': '\\leftrightarrow',
            '⇔': '\\Leftrightarrow',
        }
        
        for symbol, latex in symbols.items():
            text = text.replace(symbol, latex)
        
        return text

    
    def _convert_to_latex(self, omath):
        """Try to convert OMath to LaTeX"""
        
        try:
            # Try to get MathML first
            mathml = omath.Range.XMLNodes.Item(1).XML
            
            # Simple MathML to LaTeX conversion
            if 'mfrac' in mathml:
                # Extract numerator and denominator
                return "\\frac{num}{den}"
            
            # More conversion logic here...
            
        except:
            pass
        
        # Try to parse the linear string
        try:
            linear = omath.Range.OMaths.Item(1).LinearString
            return self._linear_to_latex(linear)
        except:
            pass
        
        return None
    
    def _linear_to_latex(self, linear_text):
        """Convert Word's linear format to LaTeX"""
        
        if not linear_text:
            return None
        
        latex = linear_text
        
        # Common conversions from Word linear to LaTeX
        conversions = {
            '≠': '\\neq',
            '≤': '\\leq',
            '≥': '\\geq',
            '∞': '\\infty',
            '∑': '\\sum',
            '∫': '\\int',
            '√': '\\sqrt',
            'α': '\\alpha',
            'β': '\\beta',
            'γ': '\\gamma',
            'π': '\\pi',
            '→': '\\rightarrow',
            '∈': '\\in',
            '∀': '\\forall',
            '∃': '\\exists',
        }
        
        for old, new in conversions.items():
            latex = latex.replace(old, new)
        
        # Handle subscripts (x_1 format)
        import re
        latex = re.sub(r'([a-zA-Z])_([0-9a-zA-Z]+)', r'\1_{\2}', latex)
        
        # Handle superscripts (x^2 format)
        latex = re.sub(r'([a-zA-Z])\^([0-9a-zA-Z]+)', r'\1^{\2}', latex)
        
        return latex
    
    def _save_equation_data(self, equations_data):
        """Save equation data to JSON file"""
        import json
        
        if self.doc and self.doc.FullName:
            output_file = Path(self.doc.FullName).parent / f"{Path(self.doc.Name).stem}_equations.json"
            
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(equations_data, f, indent=2, ensure_ascii=False)
            
            logger.info(f"Equation data saved to {output_file.name}")
    
    def _cleanup(self):
        """Clean up Word application"""
        try:
            if self.doc:
                self.doc.Close()
            if self.word:
                self.word.Quit()
        except:
            pass
        finally:
            pythoncom.CoUninitialize()

# ============= Alternative: Use Selection/Find-Replace =============
class WordEquationFinder:
    """Alternative approach using Word's Find functionality"""
    
    def __init__(self):
        pythoncom.CoInitialize()
        
    def find_and_mark_equations(self, docx_path):
        """Find equations and mark them with bookmarks"""
        
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True  # Show Word for debugging
        
        try:
            doc = word.Documents.Open(str(Path(docx_path).absolute()))
            
            # Method 1: Search for OMath
            equations_found = 0
            for i in range(1, doc.OMaths.Count + 1):
                omath = doc.OMaths.Item(i)
                
                # Select the equation
                omath.Range.Select()
                
                # Get selection
                selection = word.Selection
                
                # Add highlight
                selection.Range.HighlightColorIndex = 7  # Yellow
                
                # Add comment
                doc.Comments.Add(selection.Range, f"Equation {i}")
                
                equations_found += 1
            
            logger.info(f"Marked {equations_found} equations")
            
            # Save
            output = Path(docx_path).parent / f"{Path(docx_path).stem}_marked.docx"
            doc.SaveAs2(str(output.absolute()))
            
            return output
            
        finally:
            word.Quit()
            pythoncom.CoUninitialize()

# ============= Simple usage =============
def process_with_word_com(docx_path):
    """Process document using Word COM"""
    
    # Install pywin32 if needed
    try:
        import win32com.client
    except ImportError:
        print("Installing pywin32...")
        import subprocess
        subprocess.check_call(['pip', 'install', 'pywin32'])
        import win32com.client
    
    processor = WordCOMEquationReplacer()
    output = processor.process_document(docx_path)
    print(f"✓ Document processed: {output}")
    
    return output

# ============= Batch processor using Word COM =============
class WordCOMBatchProcessor:
    """Process multiple documents using Word COM"""
    
    def __init__(self):
        self.processor = WordCOMEquationReplacer()
    
    def process_folder(self, input_dir, output_dir=None):
        """Process all .docx files in folder"""
        
        input_dir = Path(input_dir)
        if not output_dir:
            output_dir = input_dir
        else:
            output_dir = Path(output_dir)
            output_dir.mkdir(exist_ok=True)
        
        docx_files = list(input_dir.glob("*.docx"))
        
        for docx_file in docx_files:
            try:
                print(f"Processing {docx_file.name}...")
                output = output_dir / f"{docx_file.stem}_equations_text.docx"
                self.processor.process_document(docx_file, output)
                print(f"  ✓ Saved to {output.name}")
            except Exception as e:
                print(f"  ✗ Error: {e}")

if __name__ == "__main__":
    # Install pywin32 if needed
    try:
        import win32com.client
    except ImportError:
        print("Please install pywin32:")
        print("pip install pywin32")
        exit(1)
    
    # Test with your file
    test_file = r"D:\Work 3 (20-Oct-24)\2 Side projects May 25\Encyclopedia\articles\مقالات بعد الاخراج\test\الدالة واحد لواحد (جاهزة للنشر).docx"
    
    process_with_word_com(test_file)