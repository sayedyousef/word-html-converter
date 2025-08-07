# unified_document_processor.py - FIXED VERSION
"""Orchestrator that properly handles temp files and preserves original paths."""

import logging
from pathlib import Path
import tempfile
import shutil
from utils import sanitize_filename, format_article_number
import mammoth as mammoth_lib
import re
        

class UnifiedDocumentProcessor:
    """Orchestrate existing processors for each document."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
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
            
            # Calculate output structure based on ORIGINAL document path
            relative_path = docx_file.parent.relative_to(input_folder)
            article_prefix = format_article_number(idx)
            safe_name = sanitize_filename(docx_file.stem)
            article_folder_name = f"{article_prefix}{safe_name}"
            
            # Create output folder preserving structure
            article_folder = output_folder / relative_path / article_folder_name
            article_folder.mkdir(parents=True, exist_ok=True)
            
            # Use temp file for processing
            #with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
            #    temp_latex_doc = Path(tmp.name)
            latex_path   = article_folder / f"{docx_file.stem}_latex.docx"
            anchor_path  = article_folder / f"{docx_file.stem}_anchor.docx"
            html_folder  = article_folder                               # keep as-is

            
            try:
                # Step 1: Convert Office Math to LaTeX
                print("  1. Converting Office Math to LaTeX...")
                math_converter.convert_document(docx_file, latex_path)
                
                # Step 2: Add anchors and save to output folder
                print("  2. Adding anchors...")
                anchored_path = article_folder / f"{docx_file.stem}_anchored.docx"
                anchor_registry = anchor_adder.add_anchors_to_document(latex_path, anchor_path)

                self.total_anchors += len(anchor_registry)
                
                # Save anchor registry as JSON
                import json
                anchor_json_path = article_folder / f"{docx_file.stem}.anchors.json"
                with open(anchor_json_path, 'w', encoding='utf-8') as f:
                    json.dump(anchor_registry, f, ensure_ascii=False, indent=2)
                
                # Step 3: Convert to HTML FROM THE ANCHORED DOCUMENT
                print("  3. Converting anchored document to HTML...")
                
                # Create custom converter for this document
                print(f"     [DEBUG-MAIN] Creating MammothConverter...")
                mammoth = MammothConverter()
                print(f"     [DEBUG-MAIN] MammothConverter created")
                
                print(f"     [DEBUG-MAIN] Setting css_manager...")
                mammoth.css_manager = css_manager
                print(f"     [DEBUG-MAIN] Setting use_external_css...")
                mammoth.use_external_css = True
                
                # CRITICAL: Convert from the ANCHORED document, not the temp file
                print(f"     Input: {anchored_path}")
                print(f"     Output folder: {article_folder}")
                
                print(f"     [DEBUG-MAIN] About to call _convert_document_custom...")
                print(f"     [DEBUG-MAIN] Method exists: {hasattr(self, '_convert_document_custom')}")
                
                try:
                    mammoth._convert_document(
                        anchored_path,    # The document to convert
                        article_folder,
                        output_folder,    # Where to save it
                        idx               # Index for numbering
                    )
                    '''
                    self._convert_document_custom(
                        mammoth,
                        anchored_path,       # USE THE ANCHORED DOCUMENT!
                        article_folder,      # Target output folder
                        docx_file.stem,      # Original document name
                        idx
                    )
                    '''
                    print(f"     [DEBUG-MAIN] _convert_document_custom returned successfully")
                except Exception as conv_error:
                    print(f"     [DEBUG-MAIN] ERROR calling _convert_document_custom: {conv_error}")
                    import traceback
                    print(f"     [DEBUG-MAIN] Traceback:")
                    traceback.print_exc()
                    raise
                
                # Accumulate statistics
                self.total_equations += mammoth.total_equations
                self.total_images += mammoth.total_images
                self.total_footnotes += mammoth.total_footnotes
                
                print(f"  ✓ Done: {docx_file.name}")
                print(f"     Equations: {mammoth.total_equations}, " 
                      f"Images: {mammoth.total_images}, "
                      f"Footnotes: {mammoth.total_footnotes}")
                
                # LIST ALL FILES CREATED IN THIS ARTICLE FOLDER
                print(f"  📁 Files created in {article_folder.name}:")
                for file in article_folder.iterdir():
                    if file.is_file():
                        print(f"     - {file.name}")
                    elif file.is_dir():
                        file_count = len(list(file.iterdir()))
                        print(f"     - {file.name}/ ({file_count} files)")
                
            except Exception as e:
                self.logger.error(f"Error processing {docx_file.name}: {e}", exc_info=True)
                print(f"  ✗ ERROR: {e}")
                
            #finally:
                # Clean up temp file
                #if temp_latex_doc.exists():
                #    temp_latex_doc.unlink()
        
        # Print summary
        self._print_summary(len(docx_files))
        mammoth.input_folder = anchored_path.parent
        #mammoth.output_folder = article_folder

    def _convert_document_custom(self, mammoth, docx_path: Path, 
                                 output_folder: Path, original_name: str, index: int):
        """Custom conversion that handles temp files properly."""
        html_content =""
        # Initialize mammoth's counters
        mammoth.total_equations = 0
        mammoth.total_images = 0
        mammoth.total_footnotes = 0
        mammoth.anchor_registry = {}
        
        # Setup image folder
        mammoth.current_image_folder = output_folder / "images"
        mammoth.current_image_folder.mkdir(exist_ok=True)
        mammoth.image_counter = 0
        
        # Detect equation type
        equation_type = mammoth._detect_equation_type(docx_path)
        mammoth.logger.info(f"  Detected equation type: {equation_type}")
        

        if equation_type == "office_math":
            try:
                html_content, equation_count = mammoth._convert_with_equation_markers_fixed(docx_path)
                mammoth.total_equations += equation_count
            except Exception as e:
                mammoth.logger.error(f"Error with Office Math conversion: {e}")
                print(f"     ⚠ Office Math conversion failed, using fallback")
                # Fallback to regular conversion
                with open(docx_path, "rb") as docx_file:
                    result = mammoth_lib.convert_to_html(
                        docx_file,
                        style_map=mammoth.style_map,
                        convert_image=mammoth_lib.images.img_element(mammoth._image_handler_with_anchor)
                    )
                    html_content = result.value
        else:
            # Regular conversion for LaTeX or no equations
            with open(docx_path, "rb") as docx_file:
                result = mammoth_lib.convert_to_html(
                    docx_file,
                    style_map=mammoth.style_map,
                    convert_image=mammoth_lib.images.img_element(mammoth._image_handler_with_anchor)
                )
                html_content = result.value
                
                # Process LaTeX equations if present
                if equation_type == "latex":
                    html_content = mammoth._preserve_equations_with_anchors(html_content)
                    # Count LaTeX equations
                    mammoth.total_equations = len(re.findall(
                        r'\$(?!\$)[^$\n]+?\$(?!\$)|\$\$[^$]+?\$\$|\\\[[^\]]*?\\\]|\\\([^)]*?\\\)', 
                        html_content, 
                        re.DOTALL
                    ))
        # Convert based on type
        html_path = output_folder / f"{original_name}.html"
        html_path.write_text(html_content, encoding="utf-8")
        #with open(html_path, "w", encoding="utf-8") as f:
        #    f.write(html_content, encoding="utf-8")  # or html_content if you don't need wrapper




    def _print_summary(self, doc_count: int):
        """Print processing summary."""
        print(f"\n{'='*60}")
        print(f"PROCESSING SUMMARY")
        print(f"{'='*60}")
        print(f"Documents processed: {doc_count}")
        print(f"Total equations: {self.total_equations}")
        print(f"Total images: {self.total_images}")
        print(f"Total footnotes: {self.total_footnotes}")
        print(f"Total anchors: {self.total_anchors}")
        print(f"{'='*60}", html_content)
        
        # Ensure we have HTML content
        if not html_content:
            mammoth.logger.error("No HTML content generated!")
            print("     ⚠ WARNING: No HTML content was generated")
            html_content = "<p>Error: No content could be extracted from the document.</p>"
        
        # Build complete HTML
        # Check if enhanced version exists, otherwise use regular
        if hasattr(mammoth, '_build_html_document_enhanced'):
            complete_html = mammoth._build_html_document_enhanced(
                title=original_name,
                author="",  # Can be extracted from metadata if needed
                body_html=html_content,
                has_equations=(mammoth.total_equations > 0)
            )
        elif hasattr(mammoth, '_build_html_document'):
            # Use regular method with correct parameters
            complete_html = mammoth._build_html_document(
                title=original_name,
                author="",  # Can be extracted from metadata if needed
                body_html=html_content,
                has_equations=(mammoth.total_equations > 0)
            )
        else:
            # Fallback: Build basic HTML if no build method found
            print("     ⚠ Using fallback HTML builder")
            complete_html = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <title>{original_name}</title>
</head>
<body>
    <h1>{original_name}</h1>
    {html_content}
</body>
</html>"""
        
        # If external CSS is used, modify the HTML to link to external CSS
        if mammoth.use_external_css and hasattr(mammoth, 'css_manager'):
            # Replace inline styles with link to external CSS
            css_link = '<link rel="stylesheet" href="../../../assets/css/document-styles.css">'
            complete_html = complete_html.replace('<style>', f'{css_link}\n<style>')
        
        # Save HTML
        html_path = output_folder / f"{original_name}.html"
        
        # ADD MORE LOGGING TO SEE WHAT'S HAPPENING
        print(f"  4. Saving HTML to: {html_path}")
        mammoth.logger.info(f"  Writing HTML file: {html_path}")
        
        try:
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(complete_html)
            
            # Verify the file was created
            if html_path.exists():
                file_size = html_path.stat().st_size
                print(f"  ✓ HTML saved successfully ({file_size} bytes)")
                mammoth.logger.info(f"  HTML file created: {html_path} ({file_size} bytes)")
            else:
                print(f"  ✗ ERROR: HTML file was not created!")
                mammoth.logger.error(f"  HTML file was not created at {html_path}")
                
        except Exception as e:
            print(f"  ✗ ERROR saving HTML: {e}")
            mammoth.logger.error(f"Failed to save HTML: {e}", exc_info=True)
            raise
        
        # Save anchor registry if present
        if mammoth.anchor_registry:
            import json
            registry_path = output_folder / f"{original_name}.registry.json"
            with open(registry_path, "w", encoding="utf-8") as f:
                json.dump(mammoth.anchor_registry, f, ensure_ascii=False, indent=2)
    
    def _print_summary(self, doc_count: int):
        """Print processing summary."""
        print(f"\n{'='*60}")
        print(f"PROCESSING SUMMARY")
        print(f"{'='*60}")
        print(f"Documents processed: {doc_count}")
        print(f"Total equations: {self.total_equations}")
        print(f"Total images: {self.total_images}")
        print(f"Total footnotes: {self.total_footnotes}")
        print(f"Total anchors: {self.total_anchors}")
        print(f"{'='*60}")