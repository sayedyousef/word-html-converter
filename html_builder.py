# html_builder.py
"""Build HTML output files."""

import logging
from pathlib import Path
from models import DocumentContent
from config import Config

class HTMLBuilder:
    """Handles HTML generation."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def build_html(self, content: DocumentContent, output_path: Path) -> None:
        """Build complete HTML file."""
        # Create MathJax or KaTeX script based on config
        math_script = self._get_math_script()
        
        # Build HTML
        html = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{content.title}</title>
    <meta name="author" content="{content.author}">
    {math_script}
    <style>
        body {{
            font-family: 'Arial', 'Tahoma', sans-serif;
            line-height: 1.8;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
            direction: rtl;
        }}
        h1, h2, h3, h4 {{
            color: #333;
            margin-top: 1.5em;
        }}
        .title {{
            text-align: center;
            font-size: 2em;
            margin-bottom: 0.5em;
        }}
        .author {{
            text-align: center;
            color: #666;
            margin-bottom: 2em;
        }}
        .footnote {{
            font-size: 0.9em;
            color: #666;
            border-top: 1px solid #ddd;
            margin-top: 2em;
            padding-top: 1em;
        }}
        .footnote-ref {{
            vertical-align: super;
            font-size: 0.8em;
            color: #0066cc;
            text-decoration: none;
        }}
        table {{
            border-collapse: collapse;
            width: 100%;
            margin: 1em 0;
        }}
        td, th {{
            border: 1px solid #ddd;
            padding: 8px;
            text-align: right;
        }}
        img {{
            max-width: 100%;
            height: auto;
            display: block;
            margin: 1em auto;
        }}
        .image-caption {{
            text-align: center;
            font-style: italic;
            color: #666;
            margin-top: 0.5em;
        }}
    </style>
</head>
<body>
    <h1 class="title">{content.title}</h1>
    <p class="author">{content.author}</p>
    
    <div class="content">
        {content.body_html}
    </div>
"""
        
        # Add footnotes section if present
        if content.footnotes:
            html += self._build_footnotes_html(content.footnotes)
        
        html += """
</body>
</html>"""
        
        # Write file
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(html, encoding='utf-8')
        self.logger.info(f"HTML file created: {output_path}")
    
    def _get_math_script(self) -> str:
        """Get math rendering script based on config."""
        if Config.USE_MATHJAX:
            return """
    <script>
        window.MathJax = {
            tex: {
                inlineMath: [['$', '$'], ['\\\\(', '\\\\)']],
                displayMath: [['$$', '$$'], ['\\\\[', '\\\\]']],
                processEscapes: true
            },
            options: {
                skipHtmlTags: ['script', 'noscript', 'style', 'textarea', 'pre']
            }
        };
    </script>
    <script id="MathJax-script" async src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-chtml.js"></script>
"""
        else:
            return """
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.0/dist/katex.min.css">
    <script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.0/dist/katex.min.js"></script>
    <script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.0/dist/contrib/auto-render.min.js"
        onload="renderMathInElement(document.body, {
            delimiters: [
                {left: '$$', right: '$$', display: true},
                {left: '$', right: '$', display: false},
                {left: '\\\\(', right: '\\\\)', display: false},
                {left: '\\\\[', right: '\\\\]', display: true}
            ]
        });"></script>
"""
    
    def _build_footnotes_html(self, footnotes) -> str:
        """Build HTML for footnotes section."""
        if not footnotes:
            return ""
        
        html = '\n<div class="footnotes">\n<h3>الحواشي</h3>\n<ol>\n'
        
        for footnote in footnotes:
            html += f'<li id="fn{footnote.id}">{footnote.text}</li>\n'
        
        html += '</ol>\n</div>\n'
        return html
