# logger.py
"""Logging configuration."""

import logging
import sys
from pathlib import Path
from datetime import datetime
from config import Config

def setup_logging():
    """Configure logging for the application."""
    # Create logs directory
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)
    
    # Log filename with timestamp
    log_filename = log_dir / f"word_to_html_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    
    # Create handlers
    file_handler = logging.FileHandler(log_filename, encoding='utf-8')
    file_handler.setLevel(getattr(logging, Config.LOG_LEVEL))
    file_handler.setFormatter(logging.Formatter(Config.LOG_FORMAT))
    
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(getattr(logging, Config.LOG_LEVEL))
    console_handler.setFormatter(logging.Formatter(Config.LOG_FORMAT))
    
    # Configure root logger
    logging.basicConfig(
        level=getattr(logging, Config.LOG_LEVEL),
        handlers=[file_handler, console_handler]
    )
    
    return logging.getLogger(__name__)
