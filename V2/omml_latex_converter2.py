"""
Fixed Word COM automation with multiple extraction methods
Handles errors gracefully and tries different approaches
"""

import win32com.client
import os

def extract_equations_from_word(docx_path):
    """
    Extract equations using multiple methods
    """
    
    docx_path = os.path.abspath(docx_path)
    output_file = "equations_extracted.txt"
    
    print(f"Processing: {docx_path}")
    
    # Start Word
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    
    try:
        doc = word.Documents.Open(docx_path)
        
        total_equations = doc.OMaths.Count
        print(f"Found {total_equations} equations in document\n")
        
        equations_data = []
        
        for i, equation in enumerate(doc.OMaths, 1):
            print(f"Processing equation {i}...")
            equation_info = {}
            
            # Method 1: Try to get the LinearString property
            try:
                linear = equation.LinearString
                equation_info['linear'] = linear
                print(f"  ✓ Linear format: {linear[:50]}..." if len(linear) > 50 else f"  ✓ Linear format: {linear}")
            except:
                pass
            
            # Method 2: Try Range.Text
            try:
                range_text = equation.Range.Text
                equation_info['range_text'] = range_text
                print(f"  ✓ Range text: {range_text[:50]}..." if len(range_text) > 50 else f"  ✓ Range text: {range_text}")
            except:
                pass
            
            # Method 3: Try to select and copy
            try:
                equation.Range.Select()
                selection = word.Selection
                selection_text = selection.Text
                equation_info['selection'] = selection_text
                print(f"  ✓ Selection: {selection_text[:50]}..." if len(selection_text) > 50 else f"  ✓ Selection: {selection_text}")
            except:
                pass
            
            # Method 4: Try Type property
            try:
                eq_type = equation.Type  # 0 = Inline, 1 = Display
                equation_info['type'] = "Inline" if eq_type == 0 else "Display"
                print(f"  ✓ Type: {equation_info['type']}")
            except:
                pass
            
            # Method 5: Try to get as XML (OMML)
            try:
                # Get the XML representation
                xml_range = equation.Range
                xml_text = xml_range.XML
                # Extract just the equation part if it's too long
                if '<m:oMath' in xml_text:
                    start = xml_text.find('<m:oMath')
                    end = xml_text.find('</m:oMath>') + 10
                    equation_xml = xml_text[start:end]
                    equation_info['xml'] = equation_xml[:100] + "..." if len(equation_xml) > 100 else equation_xml
                    print(f"  ✓ OMML XML captured")
            except:
                pass
            
            # Method 6: Try BuildDown/BuildUp (with better error handling)
            try:
                # Check if equation supports BuildDown
                if hasattr(equation, 'BuildDown'):
                    equation.BuildDown()
                    linear_after = equation.Range.Text
                    equation_info['builddown'] = linear_after
                    print(f"  ✓ BuildDown: {linear_after[:50]}..." if len(linear_after) > 50 else f"  ✓ BuildDown: {linear_after}")
                    equation.BuildUp()  # Restore
            except:
                pass
            
            # Method 7: Try ParentOMath
            try:
                parent = equation.ParentOMath
                if parent:
                    parent_text = parent.Range.Text
                    equation_info['parent'] = parent_text
            except:
                pass
            
            equations_data.append(equation_info)
            print()
        
        # Save all extracted data
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(f"Equations from: {os.path.basename(docx_path)}\n")
            f.write("="*70 + "\n\n")
            
            for i, eq_data in enumerate(equations_data, 1):
                f.write(f"EQUATION {i}:\n")
                f.write("-"*40 + "\n")
                
                if 'linear' in eq_data:
                    f.write(f"Linear format:\n{eq_data['linear']}\n\n")
                
                if 'range_text' in eq_data:
                    f.write(f"Range text:\n{eq_data['range_text']}\n\n")
                
                if 'selection' in eq_data:
                    f.write(f"Selection:\n{eq_data['selection']}\n\n")
                
                if 'builddown' in eq_data:
                    f.write(f"BuildDown result:\n{eq_data['builddown']}\n\n")
                
                if 'type' in eq_data:
                    f.write(f"Type: {eq_data['type']}\n\n")
                
                if 'xml' in eq_data:
                    f.write(f"OMML (partial):\n{eq_data['xml']}\n\n")
                
                f.write("="*70 + "\n\n")
        
        print(f"✅ Successfully processed {len(equations_data)} equations")
        print(f"📄 Results saved to: {output_file}")
        
        doc.Close(SaveChanges=False)
        
        return equations_data
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return []
    
    finally:
        word.Quit()


def try_alternative_com_methods(docx_path):
    """
    Try alternative COM methods for equation extraction
    """
    
    print("\nTrying alternative methods...")
    print("-"*40)
    
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    
    try:
        doc = word.Documents.Open(os.path.abspath(docx_path))
        
        # Method A: Try using Fields collection (some equations might be fields)
        print(f"Document has {doc.Fields.Count} fields")
        for i, field in enumerate(doc.Fields, 1):
            if field.Type == 3:  # wdFieldExpression
                print(f"  Field {i}: {field.Result.Text[:50]}...")
        
        # Method B: Try using Shapes collection (for embedded equations)
        print(f"Document has {doc.Shapes.Count} shapes")
        
        # Method C: Try InlineShapes (for inline equations)
        print(f"Document has {doc.InlineShapes.Count} inline shapes")
        for i, shape in enumerate(doc.InlineShapes, 1):
            if shape.Type == 3:  # wdInlineShapeEmbeddedOLEObject
                try:
                    if "Equation" in shape.OLEFormat.ProgID:
                        print(f"  Found equation object {i}")
                except:
                    pass
        
        # Method D: Direct paragraph scanning
        print(f"\nScanning {doc.Paragraphs.Count} paragraphs for equations...")
        eq_count = 0
        for para in doc.Paragraphs:
            if para.Range.OMaths.Count > 0:
                eq_count += para.Range.OMaths.Count
                text = para.Range.Text.strip()
                if text:
                    print(f"  Paragraph with equation: {text[:50]}...")
        
        print(f"Total equations found in paragraphs: {eq_count}")
        
        doc.Close(SaveChanges=False)
        
    except Exception as e:
        print(f"Error: {e}")
    
    finally:
        word.Quit()


def main():
    """
    Main function
    """
    
    print("Fixed Word COM Equation Extractor")
    print("="*70)
    print()
    
    # Your document path
    docx_path = r"D:\Work 3 (20-Oct-24)\2 Side projects May 25\Encyclopedia\articles\مقالات بعد الاخراج\test\الدالة واحد لواحد (جاهزة للنشر).docx"
    
    if not os.path.exists(docx_path):
        print(f"Error: File not found!")
        return
    
    # Try main extraction
    equations = extract_equations_from_word(docx_path)
    
    # Try alternative methods
    try_alternative_com_methods(docx_path)
    
    print("\n" + "="*70)
    print("Extraction complete! Check 'equations_extracted.txt' for results")


if __name__ == "__main__":
    main()