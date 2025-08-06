# test_conversion.py
# Minimal test script to verify the conversion is working

import sys
from pathlib import Path

# Add parent directory to path if needed
sys.path.append(str(Path(__file__).parent))

from mammoth_converter import MammothConverter
from css_manager import CSSManager
from config import Config
import logging

# Setup basic logging
logging.basicConfig(level=logging.INFO)

print("Testing document conversion...")

# Setup paths
input_folder = Path("input")  # Change this to your input folder
output_folder = Path("output")  # Change this to your output folder

# Create folders if they don't exist
input_folder.mkdir(exist_ok=True)
output_folder.mkdir(exist_ok=True)

# Check for documents
docx_files = list(input_folder.rglob("*.docx"))
docx_files = [f for f in docx_files if not f.name.startswith("~")]

if not docx_files:
    print(f"No .docx files found in {input_folder}")
    print("Please add a .docx file to test")
    sys.exit(1)

print(f"Found {len(docx_files)} documents")

# Setup CSS
css_manager = CSSManager()
css_manager.setup_css_folder()

# Initialize converter
converter = MammothConverter()
converter.css_manager = css_manager
converter.use_external_css = True

# Convert
print("Converting documents...")
converter.convert_folder(input_folder, output_folder)

print(f"Conversion complete!")
print(f"Check output in: {output_folder}")
