# ============= logger.py =============
"""Simple logging setup"""
import logging
from datetime import datetime
from config import Config

def setup_logger(name="word_converter"):
    """Setup simple logger with file and console output"""
    logger = logging.getLogger(name)
    logger.setLevel(Config.LOG_LEVEL)
    
    # Console handler - force UTF-8 encoding
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    
    # File handler - specify UTF-8 encoding
    file_handler = logging.FileHandler(Config.LOG_FILE, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    
    # Simple format
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    console.setFormatter(formatter)
    file_handler.setFormatter(formatter)
    
    logger.addHandler(console)
    logger.addHandler(file_handler)
    
    return logger