# ============= document_processor.py =============
"""Combined document processor with batch capabilities"""
from pathlib import Path
from config import Config
from logger import setup_logger
from equation_replacer import DocumentEquationReplacer
from doc_to_html_latex import DocumentToHTMLWithLatex
import json

logger = setup_logger("document_processor")

class DocumentProcessor:
    """Combined processor for single and batch document processing"""
    
    def __init__(self):
        self.input_dir = Config.INPUT_DIR
        self.output_dir = Config.OUTPUT_DIR
        self.output_format = getattr(Config, 'OUTPUT_FORMAT', 'html')  # Default to HTML
        
    def process_single_document(self, docx_path: Path, output_format: str = None) -> Path:
        """Process a single document"""
        if output_format is None:
            output_format = self.output_format
            
        docx_path = Path(docx_path)
        
        if not docx_path.exists():
            logger.error(f"File not found: {docx_path}")
            return None
        
        # Determine output path
        if output_format == 'html':
            output_path = self.output_dir / f"{docx_path.stem}.html"
            processor = DocumentToHTMLWithLatex(docx_path)
            processor.convert_to_html(output_path)
        else:
            output_path = self.output_dir / f"{docx_path.stem}_latex.docx"
            processor = DocumentEquationReplacer(docx_path)
            processor.process_document(output_path)
        
        logger.info(f"Processed {docx_path.name} -> {output_path.name}")
        return output_path
    
    def get_files(self) -> list:
        """Get all docx files from input directory"""
        files = list(self.input_dir.glob("*.docx"))
        # Filter out temp files
        files = [f for f in files if not f.name.startswith('~')]
        logger.info(f"Found {len(files)} files to process")
        return files
    
    def process_batch(self, files: list) -> list:
        """Process a batch of files"""
        results = []
        
        for docx_file in files:
            try:
                output_path = self.process_single_document(docx_file)
                results.append({
                    'input': str(docx_file),
                    'output': str(output_path),
                    'status': 'success'
                })
            except Exception as e:
                logger.error(f"Failed to process {docx_file.name}: {e}")
                results.append({
                    'input': str(docx_file),
                    'error': str(e),
                    'status': 'failed'
                })
        
        return results
    
    def save_summary(self, results: list):
        """Save processing summary"""
        summary_file = self.output_dir / 'processing_summary.json'
        with open(summary_file, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        logger.info(f"Summary saved to {summary_file}")

