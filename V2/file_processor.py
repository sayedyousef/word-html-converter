# ============= file_processor.py =============
"""File iteration and processing"""
from pathlib import Path
from typing import List, Generator
from config import Config
from logger import setup_logger
from equation_converter import EquationConverter

logger = setup_logger("file_processor")

class FileProcessor:
    """Process Word documents in batches"""

    def __init__(self):
        self.input_dir = Config.INPUT_DIR
        self.output_dir = Config.OUTPUT_DIR
        
    def get_files(self) -> List[Path]:
        """Get all supported files from input directory"""
        files = []
        for ext in Config.SUPPORTED_FORMATS:
            files.extend(self.input_dir.glob(f"*{ext}"))
        
        logger.info(f"Found {len(files)} files to process")
        return files
    
    def process_batch(self, files: List[Path]) -> None:
        """Process a batch of files"""
        for file_path in files:
            try:
                self.process_single_file(file_path)
            except Exception as e:
                logger.error(f"Failed to process {file_path.name}: {e}")
    
    def process_single_file(self, file_path: Path) -> None:
        """Process single Word document"""
        logger.info(f"Processing: {file_path.name}")
        
        # Initialize converter
        converter = EquationConverter(file_path)
        
        # Extract equations
        equations = converter.extract_equations()
        
        if not equations:
            logger.info(f"No equations found in {file_path.name}")
            return
        
        # Convert each equation
        for eq in equations:
            latex = converter.convert_to_latex(eq)
            logger.debug(f"Equation {eq['id']}: {latex}")
        
        # For now, save equation map to file
        output_file = self.output_dir / f"{file_path.stem}_equations.txt"
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(f"Equations from {file_path.name}\n")
            f.write("=" * 50 + "\n\n")
            
            for eq_id, latex in converter.equation_map.items():
                f.write(f"ID: {eq_id}\n")
                f.write(f"LaTeX: {latex}\n")
                f.write("-" * 30 + "\n")
        
        logger.info(f"Saved equations to {output_file.name}")

