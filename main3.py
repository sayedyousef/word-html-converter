# main3.py
"""Main entry point for Enhanced Word to HTML converter with equation fixes."""

import logging
import sys
from pathlib import Path
from datetime import datetime
import io

# Force UTF-8 encoding for Windows console
if sys.platform == 'win32':
    import os
    os.system('chcp 65001 >nul 2>&1')
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# Import the correct modules
from config import Config
from logger import setup_logging
from mammoth_converter import MammothConverter  # Use MammothConverter, not HTMLConverter!
from css_manager import CSSManager

def main():
    """Main function to orchestrate document conversion."""
    # Print startup message
    print("=" * 60)
    print("Enhanced Word to HTML Converter")
    print("=" * 60)
    
    # Setup logging
    print("Setting up logging...")
    logger = setup_logging()
    
    logger.info("=" * 60)
    logger.info("Enhanced Word to HTML Converter Started")
    logger.info(f"Input folder: {Config.INPUT_FOLDER}")
    logger.info(f"Output folder: {Config.OUTPUT_FOLDER}")
    logger.info(f"External CSS: {Config.USE_EXTERNAL_CSS}")
    logger.info(f"Generate anchored docs: {Config.GENERATE_ANCHORED_DOCS}")
    logger.info("=" * 60)
    
    # Setup CSS files
    print("Setting up CSS files...")
    css_manager = CSSManager(Config.CSS_FOLDER)
    css_manager.setup_css_folder()  # Ensure CSS folder exists
    
    # Copy CSS files to output folder if using external CSS
    if Config.USE_EXTERNAL_CSS:
        css_manager.copy_css_to_output(Config.OUTPUT_FOLDER)
        logger.info(f"CSS files copied to {Config.OUTPUT_FOLDER / 'assets' / 'css'}")
    else:
        logger.info("Using inline CSS (embedded in HTML)")
    
    # Validate paths
    if not Config.INPUT_FOLDER.exists():
        logger.error(f"Input folder does not exist: {Config.INPUT_FOLDER}")
        print(f"ERROR: Input folder does not exist: {Config.INPUT_FOLDER}")
        print(f"Please create the folder and add .docx files to it")
        return
    
    # Check for .docx files
    docx_files = list(Config.INPUT_FOLDER.rglob("*.docx"))
    # Filter out temporary files
    docx_files = [f for f in docx_files if not f.name.startswith("~")]
    
    if not docx_files:
        logger.warning(f"No .docx files found in {Config.INPUT_FOLDER}")
        print(f"WARNING: No .docx files found in {Config.INPUT_FOLDER}")
        print("Please add some .docx files to convert")
        return
    
    print(f"Found {len(docx_files)} document(s) to convert")
    
    # Create output folder if needed
    Config.OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    
    try:
        # Initialize converter with MammothConverter
        print("Initializing converter...")
        converter = MammothConverter()  # Use MammothConverter!
        
        # Pass CSS manager to converter
        converter.css_manager = css_manager
        converter.use_external_css = Config.USE_EXTERNAL_CSS
        
        # Process all documents
        print(f"Starting conversion of {len(docx_files)} documents...")
        logger.info(f"Processing {len(docx_files)} documents from: {Config.INPUT_FOLDER}")
        
        # Call the correct method: convert_folder not process_folder
        converter.convert_folder(Config.INPUT_FOLDER, Config.OUTPUT_FOLDER)
        
        # Optional: Generate Word documents with anchors
        if Config.GENERATE_ANCHORED_DOCS:
            print("Generating Word documents with anchors...")
            logger.info("Generating anchored Word documents...")
            
            try:
                from anchor_generator import AnchorGenerator
                anchor_gen = AnchorGenerator()
                
                html_files = list(Config.OUTPUT_FOLDER.rglob("*.html"))
                anchor_count = 0
                
                for html_file in html_files:
                    anchor_file = html_file.with_suffix('.anchors.json')
                    if anchor_file.exists():
                        try:
                            anchor_gen.create_from_html_data(html_file, anchor_file)
                            anchor_count += 1
                        except Exception as e:
                            logger.warning(f"Could not create anchored doc for {html_file}: {e}")
                
                if anchor_count > 0:
                    logger.info(f"Created {anchor_count} anchored Word documents")
                    
            except ImportError:
                logger.warning("anchor_generator.py not found - skipping anchor document generation")
                print("Note: anchor_generator.py not found - skipping anchor document generation")
        
        # Print summary
        print("\n" + "=" * 60)
        print("CONVERSION SUMMARY")
        print("=" * 60)
        print(f"✅ Documents processed: {len(docx_files)}")
        print(f"📊 Total equations: {converter.total_equations}")
        print(f"🖼️  Total images: {converter.total_images}")
        print(f"📝 Total footnotes: {converter.total_footnotes}")
        print(f"📁 Output folder: {Config.OUTPUT_FOLDER.absolute()}")
        print("=" * 60)
        
        logger.info("=" * 60)
        logger.info("Conversion completed successfully!")
        logger.info(f"Total equations processed: {converter.total_equations}")
        logger.info(f"Total images processed: {converter.total_images}")
        logger.info(f"Total footnotes processed: {converter.total_footnotes}")
        logger.info("=" * 60)
        
        print("\n✨ Conversion completed successfully!")
        
    except Exception as e:
        logger.error(f"Error during conversion: {e}", exc_info=True)
        print(f"\n❌ ERROR during conversion: {e}")
        print("Check the log file for details")
        raise

if __name__ == "__main__":
    print("Starting Enhanced Word to HTML Converter...")
    main()

