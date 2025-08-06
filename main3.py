# main3.py - COMPLETE VERSION
"""Main entry point using unified processor."""

import logging
import sys
from pathlib import Path
import io

# Force UTF-8 encoding for Windows console
if sys.platform == 'win32':
    import os
    os.system('chcp 65001 >nul 2>&1')
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

from config import Config
from logger import setup_logging
from unified_document_processor import UnifiedDocumentProcessor

def main():
    """Main function using unified processor."""
    
    # Print startup
    print("=" * 60)
    print("Document Processing System")
    print("=" * 60)
    
    # Setup logging
    print("Setting up logging...")
    logger = setup_logging()
    
    logger.info("=" * 60)
    logger.info("Document Processing Started")
    logger.info(f"Input folder: {Config.INPUT_FOLDER}")
    logger.info(f"Output folder: {Config.OUTPUT_FOLDER}")
    logger.info(f"External CSS: {Config.USE_EXTERNAL_CSS}")
    logger.info("=" * 60)
    
    # Validate paths
    if not Config.INPUT_FOLDER.exists():
        logger.error(f"Input folder does not exist: {Config.INPUT_FOLDER}")
        print(f"ERROR: Input folder does not exist: {Config.INPUT_FOLDER}")
        Config.INPUT_FOLDER.mkdir(parents=True, exist_ok=True)
        print(f"Created folder. Please add .docx files to: {Config.INPUT_FOLDER}")
        return
    
    # Check for documents
    docx_files = list(Config.INPUT_FOLDER.rglob("*.docx"))
    docx_files = [f for f in docx_files if not f.name.startswith("~")]
    
    if not docx_files:
        logger.warning(f"No .docx files found in {Config.INPUT_FOLDER}")
        print(f"WARNING: No .docx files found")
        return
    
    print(f"Found {len(docx_files)} documents to process")
    
    # Create output folder
    Config.OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    
    try:
        # Use unified processor for everything
        print("\nStarting unified document processing...")
        processor = UnifiedDocumentProcessor()
        processor.process_all_documents(Config.INPUT_FOLDER, Config.OUTPUT_FOLDER)
        
        print("\n" + "=" * 60)
        print("PROCESSING COMPLETE")
        print("=" * 60)
        print(f"✅ All documents processed")
        print(f"📁 Output folder: {Config.OUTPUT_FOLDER.absolute()}")
        print("=" * 60)
        
        logger.info("Processing completed successfully!")
        
    except Exception as e:
        logger.error(f"Error during processing: {e}", exc_info=True)
        print(f"\n❌ ERROR: {e}")
        print("Check log file for details")
        raise

if __name__ == "__main__":
    print("Starting Document Processing System...")
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nProcessing cancelled by user.")
    except Exception as e:
        print(f"\nFatal error: {e}")
        sys.exit(1)