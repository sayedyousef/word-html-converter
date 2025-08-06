# unified_document_processor.py - PROPERLY ORCHESTRATED
"""Simple orchestrator that calls existing code in order for each document."""

import logging
from pathlib import Path
import tempfile
import shutil

class UnifiedDocumentProcessor:
    """Orchestrate existing processors for each document."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        # Track statistics HERE, not in MammothConverter
        self.total_equations = 0
        self.total_images = 0
        self.total_footnotes = 0
        self.total_anchors = 0
    
    def process_all_documents(self, input_folder: Path, output_folder: Path):
        """Process each document through all steps using existing code."""
        
        # Import existing components
        from mammoth_converter import MammothConverter
        from word_anchor_adder import WordAnchorAdder
        from office_math_to_latex_converter import OfficeMathToLatexConverter
        from css_manager import CSSManager
        
        # Find all documents
        docx_files = list(input_folder.rglob("*.docx"))
        docx_files = [f for f in docx_files if not f.name.startswith("~")]
        
        print(f"Found {len(docx_files)} documents")
        
        # Setup CSS once
        css_manager = CSSManager()
        css_manager.setup_css_folder()
        css_manager.copy_css_to_output(output_folder)
        
        # Initialize components
        anchor_adder = WordAnchorAdder()
        math_converter = OfficeMathToLatexConverter()
        
        # Process each document completely
        for idx, docx_file in enumerate(docx_files, 1):
            print(f"\n[{idx}/{len(docx_files)}] Processing: {docx_file.name}")
            
            # Create NEW MammothConverter for EACH document
            mammoth = MammothConverter()
            mammoth.css_manager = css_manager
            mammoth.use_external_css = True
            mammoth.input_folder = input_folder
            mammoth.output_folder = output_folder
            
            # Initialize its counters
            mammoth.total_equations = 0
            mammoth.total_images = 0
            mammoth.total_footnotes = 0
            
            # Create output folder for this document
            safe_name = docx_file.stem.replace(' ', '_')
            article_folder = output_folder / f"article_{idx:03d}_{safe_name}"
            article_folder.mkdir(parents=True, exist_ok=True)
            
            # Use temp file for latex conversion
            with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
                temp_latex_doc = Path(tmp.name)
            
            try:
                # Step 1: Convert Office Math to LaTeX
                print("  1. Converting Office Math to LaTeX...")
                equation_count = math_converter.convert_document(docx_file, temp_latex_doc)
                
                # Step 2: Add anchors
                print("  2. Adding anchors...")
                anchored_path = article_folder / f"{docx_file.stem}_anchored.docx"
                anchor_registry = anchor_adder.add_anchors_to_document(temp_latex_doc, anchored_path)
                self.total_anchors += len(anchor_registry)
                
                # Step 3: Convert to HTML
                print("  3. Converting to HTML...")
                mammoth._convert_document(temp_latex_doc, output_folder, idx)
                
                # Accumulate statistics from this document
                self.total_equations += mammoth.total_equations
                self.total_images += mammoth.total_images
                self.total_footnotes += mammoth.total_footnotes
                
                print(f"  ✓ Done: {docx_file.name}")
                print(f"     Equations: {mammoth.total_equations}, Images: {mammoth.total_images}, Footnotes: {mammoth.total_footnotes}")
                
            finally:
                # Clean up temp file
                if temp_latex_doc.exists():
                    temp_latex_doc.unlink()
        
        # Print summary with OUR accumulated values
        print(f"\n{'='*60}")
        print(f"PROCESSING SUMMARY")
        print(f"{'='*60}")
        print(f"Documents processed: {len(docx_files)}")
        print(f"Total equations: {self.total_equations}")
        print(f"Total images: {self.total_images}")
        print(f"Total footnotes: {self.total_footnotes}")
        print(f"Total anchors: {self.total_anchors}")
        print(f"{'='*60}")