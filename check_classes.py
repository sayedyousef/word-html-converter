# check_classes.py
"""Check what classes are actually in your Python files."""

import re
from pathlib import Path

print("=" * 60)
print("CHECKING PYTHON FILES FOR CLASS DEFINITIONS")
print("=" * 60)

# Files to check
files_to_check = [
    'mammoth_converter.py',
    'integrated_converter.py',
    'enhanced_doc_processor.py',
    'html_converter.py',
    'equation_handler.py',
    'html_builder.py',
    'document_parser.py',
    'anchor_generator.py'  # May not exist yet
]

found_converters = []

for filename in files_to_check:
    file_path = Path(filename)
    if file_path.exists():
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
                
                # Find all class definitions
                classes = re.findall(r'^class\s+(\w+)', content, re.MULTILINE)
                
                if classes:
                    print(f"\n📄 {filename}:")
                    for cls in classes:
                        print(f"   ✅ class {cls}")
                        
                        # Check for key methods in converter classes
                        if 'Converter' in cls or 'converter' in filename:
                            found_converters.append((filename, cls))
                            
                            # Check for important methods
                            methods_to_check = ['convert_folder', 'process_folder', '_convert_document']
                            print(f"      Methods:")
                            for method in methods_to_check:
                                if f"def {method}" in content:
                                    print(f"         ✓ {method}()")
                else:
                    print(f"\n📄 {filename}: No classes found")
                    
        except Exception as e:
            print(f"\n📄 {filename}: Error reading - {e}")
    else:
        print(f"\n📄 {filename}: ❌ File not found")

print("\n" + "=" * 60)
print("SUMMARY - CONVERTER CLASSES FOUND:")
print("=" * 60)

if found_converters:
    for filename, classname in found_converters:
        print(f"   {filename} → {classname}")
    
    print("\n" + "=" * 60)
    print("HOW TO FIX main3.py:")
    print("=" * 60)
    
    # Suggest the correct import
    if found_converters:
        filename, classname = found_converters[0]  # Use the first one found
        module_name = filename.replace('.py', '')
        
        print(f"\nIn main3.py, change this line:")
        print(f"   from mammoth_converter import MammothConverter")
        print(f"\nTo this:")
        print(f"   from {module_name} import {classname} as MammothConverter")
        
        print(f"\nOr if you want to use the actual class name:")
        print(f"   from {module_name} import {classname}")
        print(f"   converter = {classname}()  # Instead of MammothConverter()")
else:
    print("No converter classes found!")
    print("\nYou need a file with a converter class.")
    print("The class should have these methods:")
    print("   - convert_folder() or process_folder()")
    print("   - _convert_document()")

print("\n" + "=" * 60)
print("QUICK FIX OPTIONS:")
print("=" * 60)

print("\nOption 1: Use the integrated converter")
print("   In main3.py, change import to:")
print("   from integrated_converter import IntegratedMammothConverter as MammothConverter")

print("\nOption 2: Check mammoth_converter.py")
print("   Make sure it has 'class MammothConverter:' not something else")

print("\nOption 3: Use the HTML converter")
print("   from html_converter import HTMLConverter")
print("   converter = HTMLConverter()")
print("   converter.process_folder(...)  # Note: uses process_folder not convert_folder")

print("\nRun this script to see what you actually have!")