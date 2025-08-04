# main.py
"""Main entry point for Word to HTML converter."""

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

from config import Config
from logger import setup_logging
from html_converter import HTMLConverter

def main():
    """Main function to orchestrate document conversion."""
    # Setup logging
    logger = setup_logging()
    
    logger.info("=" * 60)
    logger.info("Word to HTML Converter Started")
    logger.info("=" * 60)
    
    # Validate paths
    if not Config.INPUT_FOLDER.exists():
        logger.error(f"Input folder does not exist: {Config.INPUT_FOLDER}")
        return
    
    # Create output folder if needed
    Config.OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    
    try:
        # Initialize converter
        converter = HTMLConverter()
        
        # Process all documents
        logger.info(f"Scanning for documents in: {Config.INPUT_FOLDER}")
        converter.process_folder(Config.INPUT_FOLDER, Config.OUTPUT_FOLDER)
        
        logger.info("=" * 60)
        logger.info("Conversion completed successfully!")
        logger.info("=" * 60)
        
    except Exception as e:
        logger.error(f"Error during conversion: {e}", exc_info=True)
        raise

if __name__ == "__main__":
    main()


