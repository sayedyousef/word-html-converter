# ============= doc_to_html_latex.py =============
from pathlib import Path
from logger import setup_logger
from utils import extract_xml_from_docx, clean_latex_string, create_equation_anchor
import pypandoc
import re
logger = setup_logger("doc_to_html_latex")

class DocumentToHTMLWithLatex:
    """
    Convert document to HTML with equations as LaTeX text
    This gives more control over the output format
    """
    
    def __init__(self, docx_path: Path):
        self.docx_path = Path(docx_path)
        self.equations = []
        
    def convert_to_html(self, output_path: Path = None) -> str:
        """Convert document to HTML with LaTeX equations as text"""
        
        if not output_path:
            output_path = self.docx_path.parent / f"{self.docx_path.stem}.html"
        
        logger.info(f"Converting {self.docx_path.name} to HTML")
        
        # Convert to HTML with math preserved
        html = pypandoc.convert_file(
            str(self.docx_path),
            'html5',
            format='docx',
            extra_args=[
                '--mathjax',  # Preserve math
                '--standalone',
                '--self-contained',
                '--extract-media=images',
            ]
        )
        
        # Process equations - replace MathJax with plain LaTeX
        html_processed = self._replace_mathjax_with_latex(html)
        
        # Save HTML
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_processed)
        
        logger.info(f"HTML saved to {output_path.name}")
        return html_processed
    
    def _replace_mathjax_with_latex(self, html: str) -> str:
        """Replace MathJax equations with plain LaTeX text"""
        
        # Pattern for MathJax inline: \(...\)
        html = re.sub(
            r'\\\((.*?)\\\)',
            lambda m: f'<span class="latex-inline" data-latex="{m.group(1)}">{m.group(1)}</span>',
            html,
            flags=re.DOTALL
        )
        
        # Pattern for MathJax display: \[...\]
        html = re.sub(
            r'\\\[(.*?)\\\]',
            lambda m: f'<div class="latex-display" data-latex="{m.group(1)}">{m.group(1)}</div>',
            html,
            flags=re.DOTALL
        )
        
        # Remove MathJax script tags
        html = re.sub(r'<script[^>]*mathjax[^>]*>.*?</script>', '', html, flags=re.DOTALL | re.IGNORECASE)
        
        # Add custom CSS for equation styling
        css = """
        <style>
            .latex-inline {
                font-family: 'Courier New', monospace;
                background-color: #f0f0f0;
                padding: 2px 4px;
                border-radius: 3px;
                white-space: nowrap;
            }
            .latex-display {
                font-family: 'Courier New', monospace;
                background-color: #f0f0f0;
                padding: 10px;
                margin: 10px 0;
                border-radius: 5px;
                overflow-x: auto;
                white-space: pre;
            }
        </style>
        """
        
        # Insert CSS before </head>
        html = html.replace('</head>', css + '</head>')
        
        return html

