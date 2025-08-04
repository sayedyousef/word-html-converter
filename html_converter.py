# html_converter.py
"""Main HTML converter orchestrator."""

import logging
import shutil
from pathlib import Path
from typing import List
from document_parser import DocumentParser
from html_builder import HTMLBuilder
from equation_handler import EquationHandler
from utils import sanitize_filename, format_article_number
from models import DocumentContent

class HTMLConverter:
    """Main converter class that orchestrates the conversion process."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.parser = DocumentParser()
        self.builder = HTMLBuilder()
        self.equation_handler = EquationHandler()
        self.document_counter = 0
        
    def process_folder(self, input_folder: Path, output_folder: Path) -> None:
        """Process all Word documents in folder and subfolders."""
        # Store input folder for relative path calculation
        self.input_folder = input_folder
        
        # Find all .docx files
        docx_files = list(input_folder.rglob("*.docx"))
        
        # Filter out temporary files
        docx_files = [f for f in docx_files if not f.name.startswith("~")]
        
        self.logger.info(f"Found {len(docx_files)} documents to process")
        
        # Process each file
        for idx, docx_file in enumerate(docx_files, 1):
            self._process_single_document(docx_file, output_folder, idx)

    def _process_single_document(self, docx_path: Path, output_base: Path, index: int) -> None:
        """Process a single document."""
        try:
            self.logger.info(f"Processing [{index}]: {docx_path.name}")
            
            # Parse document
            content = self.parser.parse_document(docx_path)
            if not content:
                self.logger.error(f"Failed to parse: {docx_path}")
                return
            
            # Process equations if present
            if content.has_equations:
                content.body_html = self.equation_handler.process_equations(content.body_html)
            
            # Get relative path from input folder
            relative_path = docx_path.parent.relative_to(self.input_folder)
            
            # Create output structure preserving folder hierarchy
            article_prefix = format_article_number(index)
            safe_name = sanitize_filename(docx_path.stem)
            article_folder_name = f"{article_prefix}{safe_name}"
            
            # Create article folder with preserved path structure
            article_folder = output_base / relative_path / article_folder_name
            article_folder.mkdir(parents=True, exist_ok=True)
            
            # Create images subfolder if document has images
            if content.images:
                images_folder = article_folder / "images"
                images_folder.mkdir(exist_ok=True)
                self._extract_images(docx_path, images_folder, content.images)
            
            # Build and save HTML
            html_path = article_folder / f"{safe_name}.html"
            self.builder.build_html(content, html_path)
            
            self.logger.info(f"Converted: {docx_path.name} → {html_path}")
            
        except Exception as e:
            self.logger.error(f"Error processing {docx_path}: {e}", exc_info=True)

    def _extract_images(self, docx_path: Path, images_folder: Path, image_infos: List) -> None:
        """Extract images from document."""
        try:
            # This is a simplified version - in production you'd want to
            # properly extract and save images from the docx
            self.logger.info(f"Document contains {len(image_infos)} images")
        except Exception as e:
            self.logger.warning(f"Could not extract images: {e}")
