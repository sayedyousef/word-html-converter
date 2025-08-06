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

