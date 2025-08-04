# ===================================

# main_mammoth.py
"""Main entry point for mammoth converter."""

import logging
import sys
from pathlib import Path
import io

# Force UTF-8 encoding
if sys.platform == 'win32':
    import os
    os.system('chcp 65001 >nul 2>&1')
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

from config import Config
from logger import setup_logging
from mammoth_converter import MammothConverter

def main():
    """Main function."""
    # Setup logging
    logger = setup_logging()
    
    logger.info("=" * 60)
    logger.info("Mammoth Word to HTML Converter")
    logger.info("=" * 60)
    
    # Validate paths
    if not Config.INPUT_FOLDER.exists():
        logger.error(f"Input folder not found: {Config.INPUT_FOLDER}")
        return
    
    # Create output folder
    Config.OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    
    try:
        # Convert
        converter = MammothConverter()
        converter.convert_folder(Config.INPUT_FOLDER, Config.OUTPUT_FOLDER)
        
        logger.info("=" * 60)
        logger.info("Conversion completed!")
        logger.info("=" * 60)
        
    except Exception as e:
        logger.error(f"Error: {e}", exc_info=True)

if __name__ == "__main__":
    main()