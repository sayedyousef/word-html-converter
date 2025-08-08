# ============= main2.py =============
"""Main entry point using Config"""
from config import Config
from logger import setup_logger
from document_processor import DocumentProcessor
import os
from pathlib import Path

# Point to your extracted pandoc
pandoc_exe = Path(__file__).parent / "pandoc-3.7.0.2" / "pandoc.exe"
if pandoc_exe.exists():
    os.environ['PATH'] = str(pandoc_exe.parent) + ";" + os.environ.get('PATH', '')
    print(f"✓ Using pandoc from: {pandoc_exe}")
else:
    print(f"✗ Pandoc not found at: {pandoc_exe}")

logger = setup_logger("main")

def main():
    """Main processing function"""
    logger.info("Starting Word to HTML/LaTeX Converter")
    
    # Ensure directories exist
    Config.ensure_directories()
    
    # Initialize processor
    processor = DocumentProcessor()
    
    # Get all files
    files = processor.get_files()
    
    if not files:
        logger.warning("No files found to process")
        return
    
    # Process in batches
    batch_size = Config.BATCH_SIZE
    all_results = []
    
    for i in range(0, len(files), batch_size):
        batch = files[i:i + batch_size]
        logger.info(f"Processing batch {i//batch_size + 1} ({len(batch)} files)")
        
        # Process this batch
        batch_results = processor.process_batch(batch)
        all_results.extend(batch_results)
    
    # Save overall summary
    processor.save_summary(all_results)
    
    # Print summary
    successful = len([r for r in all_results if r['status'] == 'success'])
    failed = len([r for r in all_results if r['status'] == 'failed'])
    
    logger.info(f"Processing complete: {successful} successful, {failed} failed")

if __name__ == "__main__":
    main()

