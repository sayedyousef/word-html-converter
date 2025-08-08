# ============= word_equation_processor.py =============
"""Complete Word COM equation processor - single file solution"""

import win32com.client
from pathlib import Path
import pythoncom
import json
import re

class WordEquationProcessor:
    """Process Word documents to replace equations with LaTeX text"""
    
    def __init__(self):
        pythoncom.CoInitialize()
        self.word = None
        self.doc = None
        
    def process_document(self, docx_path):
        """Main processing function"""
        
        docx_path = Path(docx_path).absolute()
        output_path = docx_path.parent / f"{docx_path.stem}_equations_text.docx"
        json_path = docx_path.parent / f"{docx_path.stem}_equations.json"
        
        print(f"\n📁 Processing: {docx_path.name}")
        print(f"📁 Output: {output_path.name}\n")
        
        try:
            # Start Word
            print("Starting Word...")
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = False
            
            # Open document
            print(f"Opening document...")
            self.doc = self.word.Documents.Open(str(docx_path))
            
            # Process equations
            equations_data = self._process_all_equations()
            
            # Save document
            print(f"Saving modified document...")
            self.doc.SaveAs2(str(output_path))
            
            # Save JSON data
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(equations_data, f, indent=2, ensure_ascii=False)
            
            print(f"\n✅ SUCCESS!")
            print(f"   📄 Word: {output_path}")
            print(f"   📋 JSON: {json_path}")
            print(f"   ✓ Processed {len(equations_data)} equations")
            
            return output_path
            
        except Exception as e:
            print(f"❌ Error: {e}")
            raise
            
        finally:
            if self.doc:
                self.doc.Close()
            if self.word:
                self.word.Quit()
            pythoncom.CoUninitialize()
    
    def _process_all_equations(self):
        """Process all equations in document"""
        
        equations_data = []
        omaths = self.doc.OMaths
        total = omaths.Count
        
        print(f"Found {total} equations")
        
        # Process from end to start to avoid index issues
        for i in range(total, 0, -1):
            try:
                omath = omaths.Item(i)
                
                # Get equation text
                latex_text = self._extract_equation_text(omath)
                
                # Replace with text
                equation_range = omath.Range
                
                # Convert to normal text first
                try:
                    omath.ConvertToNormalText()
                except:
                    # If conversion fails, just replace the range
                    equation_range.Text = latex_text
                
                # Format as equation-like
                equation_range.Font.Name = "Courier New"
                equation_range.Font.Size = equation_range.Font.Size
                equation_range.Shading.BackgroundPatternColor = 0xF0F0F0  # Light gray
                
                # Add bookmark
                bookmark_name = f"eq_{i}"
                try:
                    self.doc.Bookmarks.Add(bookmark_name, equation_range)
                except:
                    pass
                
                equations_data.append({
                    'index': i,
                    'latex': latex_text,
                    'bookmark': bookmark_name
                })
                
                print(f"  ✓ Equation {i}: {latex_text[:50]}...")
                
            except Exception as e:
                print(f"  ⚠ Equation {i}: Failed - {str(e)[:50]}")
                
                # Try fallback method
                try:
                    latex_text = self._fallback_extraction(i)
                    equations_data.append({
                        'index': i,
                        'latex': latex_text,
                        'bookmark': f"eq_{i}"
                    })
                except:
                    pass
        
        return equations_data
    
    def _extract_equation_text(self, omath):
        """Extract clean LaTeX from equation"""
        
        # Try LinearString first (best option)
        try:
            text = omath.LinearString
            if text:
                return self._clean_equation_text(text)
        except:
            pass
        
        # Try Range.Text
        try:
            text = omath.Range.Text
            if text:
                return self._clean_equation_text(text)
        except:
            pass
        
        # Default
        return "[equation]"
    
    def _fallback_extraction(self, index):
        """Fallback for problematic equations"""
        
        omath = self.doc.OMaths.Item(index)
        
        # Select and copy
        omath.Range.Select()
        self.word.Selection.Copy()
        
        # Delete equation
        self.word.Selection.Delete()
        
        # Paste as text
        self.word.Selection.PasteSpecial(DataType=2)  # wdPasteText
        
        # Get text
        text = self.word.Selection.Text
        
        # Format
        self.word.Selection.Font.Name = "Courier New"
        self.word.Selection.Shading.BackgroundPatternColor = 0xF0F0F0
        
        return self._clean_equation_text(text)
    
    def _clean_equation_text(self, text):
        """Clean and convert equation text to LaTeX"""
        
        if not text:
            return "[equation]"
        
        # Remove control characters
        text = text.replace('\r', ' ')
        text = text.replace('\n', ' ')
        text = text.replace('\x07', '')
        text = text.replace('\x0b', '')
        text = text.replace('\t', ' ')
        
        # Fix Unicode subscripts
        subscripts = {
            '₀': '_0', '₁': '_1', '₂': '_2', '₃': '_3', '₄': '_4',
            '₅': '_5', '₆': '_6', '₇': '_7', '₈': '_8', '₉': '_9',
        }
        for char, replacement in subscripts.items():
            text = text.replace(char, replacement)
        
        # Fix Unicode superscripts
        superscripts = {
            '⁰': '^0', '¹': '^1', '²': '^2', '³': '^3', '⁴': '^4',
            '⁵': '^5', '⁶': '^6', '⁷': '^7', '⁸': '^8', '⁹': '^9',
        }
        for char, replacement in superscripts.items():
            text = text.replace(char, replacement)
        
        # Convert symbols to LaTeX
        symbols = {
            '≠': '\\neq',
            '≤': '\\leq',
            '≥': '\\geq',
            '∞': '\\infty',
            '∑': '\\sum',
            '∫': '\\int',
            '√': '\\sqrt',
            'α': '\\alpha',
            'β': '\\beta',
            'π': '\\pi',
            '×': '\\times',
            '÷': '\\div',
            '→': '\\rightarrow',
            '∈': '\\in',
        }
        for symbol, latex in symbols.items():
            text = text.replace(symbol, latex)
        
        # Fix patterns like x1 to x_1
        text = re.sub(r'([a-zA-Z])([0-9]+)', r'\1_{\2}', text)
        
        # Clean multiple spaces
        text = ' '.join(text.split())
        
        return text.strip()

def process_folder(input_folder):
    """Process all .docx files in folder"""
    
    input_path = Path(input_folder)
    processor = WordEquationProcessor()
    
    docx_files = list(input_path.glob("*.docx"))
    # Skip temporary files
    docx_files = [f for f in docx_files if not f.name.startswith('~')]
    
    print(f"Found {len(docx_files)} Word documents\n")
    
    for docx_file in docx_files:
        try:
            processor.process_document(docx_file)
            print("-" * 50)
        except Exception as e:
            print(f"Failed: {e}")
            print("-" * 50)

if __name__ == "__main__":
    # Your working path
    input_folder = r"D:\Work 3 (20-Oct-24)\2 Side projects May 25\Encyclopedia\articles\مقالات بعد الاخراج\test"
    
    # Process single file or whole folder
    single_file = Path(input_folder) / "الدالة واحد لواحد (جاهزة للنشر).docx"
    
    if single_file.exists():
        processor = WordEquationProcessor()
        processor.process_document(single_file)
    else:
        process_folder(input_folder)