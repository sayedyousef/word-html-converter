# config.py
"""Configuration settings for Word to HTML converter."""

from pathlib import Path

class Config:
    """Application configuration."""
    
    # Paths
    #INPUT_FOLDER = Path(r"D:\Work 3 (20-Oct-24)\2 Side projects May 25\Encyclopedia\articles\مقالات بعد الاخراج")
    #INPUT_FOLDER = Path(r"D:\Work 3 (20-Oct-24)\2 Side projects May 25\Encyclopedia\articles\مقالات بعد الاخراج\الرياضيات والفيزياء")
    INPUT_FOLDER = Path(r"D:\Work 3 (20-Oct-24)\2 Side projects May 25\Encyclopedia\articles\مقالات بعد الاخراج\test")
    OUTPUT_FOLDER = Path(r"D:\Work 3 (20-Oct-24)\2 Side projects May 25\Encyclopedia\articles\html")
    
    # Processing settings
    IMAGE_EXTENSIONS = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.svg']
    
    # LaTeX detection patterns
    LATEX_INLINE_DELIMITERS = [
        (r'\$', r'\$'),           # $...$
        (r'\\(', r'\\)'),         # \(...\)
    ]
    LATEX_DISPLAY_DELIMITERS = [
        (r'\$\$', r'\$\$'),       # $$...$$
        (r'\\[', r'\\]'),         # \[...\]
    ]
    
    # HTML template settings
    USE_MATHJAX = True  # True for MathJax, False for KaTeX
    
    # Logging
    LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    LOG_LEVEL = 'INFO'

    USE_EXTERNAL_CSS = True
    CSS_FOLDER = Path("assets/css")
    GENERATE_ANCHORED_DOCS = True  # Set to True if you want Word docs with anchors
    CREATE_ANCHORED_WORD_DOCS = True  # Create Word docs with anchors


