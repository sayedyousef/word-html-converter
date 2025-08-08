# ============= config.py =============
"""Configuration for Word Equation Converter"""
import os
from pathlib import Path

class Config:
    # Paths
    INPUT_DIR = Path(r"D:\Work 3 (20-Oct-24)\2 Side projects May 25\Encyclopedia\articles\مقالات بعد الاخراج\test")
    OUTPUT_DIR = Path(r"D:\Work 3 (20-Oct-24)\2 Side projects May 25\Encyclopedia\articles\html2")
    TEMP_DIR = Path(r"D:\Work 3 (20-Oct-24)\2 Side projects May 25\Encyclopedia\articles\tmp")
    
    # Processing
    SUPPORTED_FORMATS = ['.docx']
    BATCH_SIZE = 10
    #OUTPUT_FORMAT = 'html'  # 'html' or 'docx'
    OUTPUT_FORMAT = 'docx'
    
    # Equation Processing
    EQUATION_ANCHOR_PREFIX = "eq_"
    EQUATION_MARKER = "[[EQUATION]]"
    
    # Logging
    LOG_LEVEL = "INFO"
    LOG_FILE = "conversion.log"
    
    @classmethod
    def ensure_directories(cls):
        """Create necessary directories if they don't exist"""
        cls.INPUT_DIR.mkdir(parents=True, exist_ok=True)
        cls.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        cls.TEMP_DIR.mkdir(parents=True, exist_ok=True)