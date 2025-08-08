# ============= main.py =============
"""Main entry point"""
from config import Config
from logger import setup_logger
from file_processor import FileProcessor

logger = setup_logger("main")

def main():
    """Main processing function"""
    logger.info("Starting Word Equation Converter")
    
    # Ensure directories exist
    Config.ensure_directories()
    # Initialize processor
    processor = FileProcessor()
    
    # Get files
    files = processor.get_files()
    
    if not files:
        logger.warning("No files found to process")
        return
    
    # Process in batches
    batch_size = Config.BATCH_SIZE
    for i in range(0, len(files), batch_size):
        batch = files[i:i + batch_size]
        logger.info(f"Processing batch {i//batch_size + 1}")
        processor.process_batch(batch)

    logger.info("Processing complete")

if __name__ == "__main__":
    main()