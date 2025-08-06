# css_manager.py
"""CSS Manager for handling external stylesheets in the document processing project."""

from pathlib import Path
from typing import List, Optional
import shutil
import logging

class CSSManager:
    """Manages CSS files for the document conversion project."""
    
    def __init__(self, css_folder: Path = None):
        """Initialize CSS manager with folder path."""
        self.logger = logging.getLogger(__name__)
        self.css_folder = css_folder or Path("assets/css")
        self.css_files = {
            'base': 'base-styles.css',
            'equations': 'equation-styles.css',
            'tables': 'table-styles.css',
            'images': 'image-styles.css',
            'footnotes': 'footnote-styles.css',
            'anchors': 'anchor-styles.css',
            'print': 'print-styles.css',
            'responsive': 'responsive-styles.css',
            'theme': 'theme-styles.css',
            'utilities': 'utilities.css'
        }
        
    def setup_css_folder(self):
        """Create CSS folder structure and copy CSS files."""
        self.css_folder.mkdir(parents=True, exist_ok=True)
        
        # Create all CSS files if they don't exist
        for css_name, css_file in self.css_files.items():
            css_path = self.css_folder / css_file
            if not css_path.exists():
                self.logger.info(f"Creating CSS file: {css_file}")
                # Here you would copy the actual CSS content
                # For now, we'll just create a placeholder
                css_path.touch()
    
    def get_css_links(self, css_types: List[str] = None, relative_path: str = "") -> str:
        """Generate HTML link tags for CSS files."""
        if css_types is None:
            css_types = ['base', 'equations', 'tables', 'images', 'footnotes', 'anchors', 'print', 'responsive']
        
        links = []
        for css_type in css_types:
            if css_type in self.css_files:
                css_file = self.css_files[css_type]
                css_path = f"{relative_path}assets/css/{css_file}"
                links.append(f'    <link rel="stylesheet" href="{css_path}">')
        
        return '\n'.join(links)
    
    def get_inline_css(self, css_types: List[str] = None) -> str:
        """Get CSS content as inline styles (fallback option)."""
        if css_types is None:
            css_types = ['base', 'equations', 'tables', 'images', 'footnotes']
        
        css_content = []
        for css_type in css_types:
            if css_type in self.css_files:
                css_file = self.css_files[css_type]
                css_path = self.css_folder / css_file
                if css_path.exists():
                    with open(css_path, 'r', encoding='utf-8') as f:
                        css_content.append(f"/* {css_file} */\n{f.read()}")
        
        return '\n'.join(css_content)
    
    def copy_css_to_output(self, output_folder: Path):
        """Copy CSS files to output folder for distribution."""
        css_output = output_folder / "assets" / "css"
        css_output.mkdir(parents=True, exist_ok=True)
        
        for css_file in self.css_files.values():
            src = self.css_folder / css_file
            dst = css_output / css_file
            if src.exists():
                shutil.copy2(src, dst)
                self.logger.debug(f"Copied {css_file} to output folder")


# Updated mammoth_converter.py methods
"""
Update these methods in your mammoth_converter.py to use external CSS
"""

def _build_html_document_with_external_css(self, title, author, body_html, has_equations, 
                                           use_external_css=True, css_manager=None):
    """Build HTML document with external CSS files."""
    
    # Initialize CSS manager if not provided
    if css_manager is None:
        css_manager = CSSManager()
    
    # Determine which CSS files to include
    css_types = ['base', 'tables', 'images', 'footnotes', 'responsive', 'print', 'utilities']
    if has_equations:
        css_types.insert(1, 'equations')
    if self.anchor_registry:
        css_types.insert(2, 'anchors')
    
    # Generate CSS links or inline styles
    if use_external_css:
        # Calculate relative path from HTML file to CSS folder
        # Assuming HTML is in output/article_folder/file.html
        # and CSS is in output/assets/css/
        relative_path = "../../"
        css_section = css_manager.get_css_links(css_types, relative_path)
    else:
        # Fallback to inline CSS
        css_content = css_manager.get_inline_css(css_types)
        css_section = f"    <style>\n{css_content}\n    </style>"
    
    # MathJax configuration
    math_script = ""
    if has_equations:
        math_script = """
    <!-- MathJax for equations -->
    <script>
        window.MathJax = {
            tex: {
                inlineMath: [['$', '$'], ['\\\\(', '\\\\)']],
                displayMath: [['$$', '$$'], ['\\\\[', '\\\\]']],
                processEscapes: true,
                processEnvironments: true
            },
            svg: {
                fontCache: 'global',
                scale: 1.1
            }
        };
    </script>
    <script id="MathJax-script" async src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-svg.js"></script>
"""
    
    return f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <meta name="author" content="{author}">
    <meta name="generator" content="Document Processing System">
    
    <!-- External CSS Files -->
{css_section}
    
    {math_script}
</head>
<body>
    <div class="content">
        <h1 class="title">{title}</h1>
        <p class="author">{author}</p>
        
        {body_html}
    </div>
    
    <!-- JavaScript enhancements -->
    <script>
        // Equation numbering
        document.addEventListener('DOMContentLoaded', function() {{
            const displayEquations = document.querySelectorAll('.display-math');
            displayEquations.forEach((eq, index) => {{
                if (!eq.querySelector('.equation-number')) {{
                    const number = document.createElement('span');
                    number.className = 'equation-number';
                    number.textContent = `(${{index + 1}})`;
                    eq.appendChild(number);
                }}
            }});
            
            // Smooth scroll to anchors
            if (window.location.hash) {{
                setTimeout(() => {{
                    const target = document.querySelector(window.location.hash);
                    if (target) {{
                        target.scrollIntoView({{ behavior: 'smooth', block: 'center' }});
                    }}
                }}, 500);
            }}
        }});
    </script>
</body>
</html>"""


# config.py additions
"""
Add these settings to your config.py file
"""

class Config:
    # ... existing config ...
    
    # CSS Settings
    USE_EXTERNAL_CSS = True  # Set to False to use inline CSS
    CSS_FOLDER = Path("assets/css")  # Folder containing CSS files
    
    # CSS files to include for different document types
    CSS_PROFILES = {
        'standard': ['base', 'tables', 'images', 'footnotes', 'responsive', 'print'],
        'math': ['base', 'equations', 'tables', 'images', 'footnotes', 'anchors', 'responsive', 'print'],
        'simple': ['base', 'responsive'],
        'full': ['base', 'equations', 'tables', 'images', 'footnotes', 'anchors', 
                 'responsive', 'print', 'theme', 'utilities']
    }
    
    # Theme settings
    ENABLE_DARK_MODE = True  # Add auto-dark class to body
    ENABLE_HIGH_CONTRAST = False  # Option for accessibility


# main2.py updates
"""
Update your main2.py to setup CSS files
"""

import logging
import sys
from pathlib import Path
from datetime import datetime
import io

# Force UTF-8 encoding for Windows console
if sys.platform == 'win32':
    import os
    os.system('chcp 65001 >nul 2>&1')
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

from config import Config
from logger import setup_logging
from mammoth_converter import MammothConverter
from css_manager import CSSManager  # New import

def main():
    """Main function to orchestrate document conversion."""
    # Setup logging
    logger = setup_logging()
    
    logger.info("=" * 60)
    logger.info("Enhanced Word to HTML Converter")
    logger.info("=" * 60)
    
    # Setup CSS files
    css_manager = CSSManager(Config.CSS_FOLDER)
    css_manager.setup_css_folder()
    
    # Copy CSS files to output folder
    css_manager.copy_css_to_output(Config.OUTPUT_FOLDER)
    logger.info(f"CSS files copied to {Config.OUTPUT_FOLDER / 'assets' / 'css'}")
    
    # Validate paths
    if not Config.INPUT_FOLDER.exists():
        logger.error(f"Input folder does not exist: {Config.INPUT_FOLDER}")
        return
    
    # Create output folder if needed
    Config.OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    
    try:
        # Initialize converter with CSS manager
        converter = MammothConverter()
        converter.css_manager = css_manager  # Pass CSS manager to converter
        converter.use_external_css = Config.USE_EXTERNAL_CSS
        
        # Process all documents
        logger.info(f"Scanning for documents in: {Config.INPUT_FOLDER}")
        converter.convert_folder(Config.INPUT_FOLDER, Config.OUTPUT_FOLDER)
        
        logger.info("=" * 60)
        logger.info("Conversion completed successfully!")
        logger.info(f"Total equations processed: {converter.total_equations}")
        logger.info(f"Total images processed: {converter.total_images}")
        logger.info("=" * 60)
        
    except Exception as e:
        logger.error(f"Error during conversion: {e}", exc_info=True)
        raise

if __name__ == "__main__":
    main()


# setup_css_files.py
"""
Standalone script to create all CSS files with content
Run this once to set up your CSS files
"""

from pathlib import Path
import logging

def create_css_files():
    """Create all CSS files with their content."""
    
    css_folder = Path("assets/css")
    css_folder.mkdir(parents=True, exist_ok=True)
    
    # Define CSS content for each file
    css_content = {
        'base-styles.css': """/* Base styles - copy from the artifact above */""",
        'equation-styles.css': """/* Equation styles - copy from the artifact above */""",
        'table-styles.css': """/* Table styles - copy from the artifact above */""",
        'image-styles.css': """/* Image styles - copy from the artifact above */""",
        'footnote-styles.css': """/* Footnote styles - copy from the artifact above */""",
        'anchor-styles.css': """/* Anchor styles - copy from the artifact above */""",
        'print-styles.css': """/* Print styles - copy from the artifact above */""",
        'responsive-styles.css': """/* Responsive styles - copy from the artifact above */""",
        'theme-styles.css': """/* Theme styles - copy from the artifact above */""",
        'utilities.css': """/* Utility classes - copy from the artifact above */"""
    }
    
    # Create each CSS file
    for filename, content in css_content.items():
        filepath = css_folder / filename
        # Note: You need to replace the placeholder with actual CSS content from the first artifact
        filepath.write_text(content, encoding='utf-8')
        print(f"Created: {filepath}")
    
    print(f"\nAll CSS files created in {css_folder}")
    print("Remember to copy the actual CSS content from the provided styles!")

if __name__ == "__main__":
    create_css_files()
