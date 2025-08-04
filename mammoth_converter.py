# mammoth_converter.py
"""Minimal Word to HTML converter using mammoth."""

import mammoth
import logging
import re
import docx
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from utils import sanitize_filename, format_article_number, detect_latex_equations

class MammothConverter:
    """Enhanced converter using mammoth with all features."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        
        # Enhanced style map with more mappings
        self.style_map = """
        p[style-name='Heading 1'] => h1:fresh
        p[style-name='Heading 2'] => h2:fresh
        p[style-name='Heading 3'] => h3:fresh
        p[style-name='Heading 4'] => h4:fresh
        p[style-name='Title'] => h1.title:fresh
        p[style-name='Subtitle'] => h2.subtitle:fresh
        p[style-name='Quote'] => blockquote:fresh
        p[style-name='Caption'] => p.caption:fresh
        r[style-name='Strong'] => strong
        r[style-name='Emphasis'] => em
        """
    
    def convert_folder(self, input_folder: Path, output_folder: Path):
        """Convert all Word documents in folder."""
        self.input_folder = input_folder
        
        # Find all .docx files
        docx_files = list(input_folder.rglob("*.docx"))
        docx_files = [f for f in docx_files if not f.name.startswith("~")]
        
        self.logger.info(f"Found {len(docx_files)} documents")
        
        # Track statistics
        self.total_equations = 0
        self.total_footnotes = 0
        self.total_images = 0
        
        for idx, docx_file in enumerate(docx_files, 1):
            self._convert_document(docx_file, output_folder, idx)
        
        # Log summary
        self.logger.info("=" * 60)
        self.logger.info("Conversion Summary:")
        self.logger.info(f"  Total equations found: {self.total_equations}")
        self.logger.info(f"  Total footnotes: {self.total_footnotes}")
        self.logger.info(f"  Total images: {self.total_images}")
        self.logger.info("=" * 60)
    
    def _convert_document(self, docx_path: Path, output_base: Path, index: int):
        """Convert single document with enhanced features."""
        try:
            self.logger.info(f"Converting [{index}]: {docx_path.name}")
            
            # Extract metadata using python-docx
            metadata = self._extract_metadata(docx_path)
            
            # Get relative path to preserve folder structure
            relative_path = docx_path.parent.relative_to(self.input_folder)
            
            # Create output structure
            article_prefix = format_article_number(index)
            safe_name = sanitize_filename(docx_path.stem)
            article_folder_name = f"{article_prefix}{safe_name}"
            
            # Create output folder with preserved structure
            article_folder = output_base / relative_path / article_folder_name
            article_folder.mkdir(parents=True, exist_ok=True)
            
            # Setup image handling
            self.current_image_folder = article_folder / "images"
            self.current_image_folder.mkdir(exist_ok=True)
            self.image_counter = 0
            
            # Detect equation type
            equation_type = self._detect_equation_type(docx_path)
            self.logger.info(f"  Detected equation type: {equation_type}")
            
            # Handle Office Math equations with position preservation
            if equation_type == "office_math":
                html_content, equation_count = self._convert_with_equation_markers(docx_path)
                self.total_equations += equation_count
                has_equations = equation_count > 0
            else:
                # Regular conversion for LaTeX or no equations
                with open(docx_path, "rb") as docx_file:
                    result = mammoth.convert_to_html(
                        docx_file,
                        style_map=self.style_map,
                        convert_image=mammoth.images.img_element(self._image_handler)
                    )
                
                html_content = result.value
                
                # Check for warnings (but suppress oMath warnings)
                if result.messages:
                    for msg in result.messages:
                        if "oMath" not in str(msg):
                            self.logger.warning(f"{docx_path.name}: {msg}")
                
                # Process LaTeX equations if present
                if equation_type == "latex":
                    html_content = self._preserve_equations(html_content)
                
                # Check for equations
                has_equations, found_equations = detect_latex_equations(html_content)
                if has_equations:
                    self.total_equations += len(found_equations)
                    self.logger.info(f"  Found {len(found_equations)} LaTeX equations")
            
            # Process other content
            html_content = self._add_footnote_backlinks(html_content)
            html_content = self._enhance_tables(html_content)
            
            # Check footnote count
            footnote_count = html_content.count('<li id="fn-')
            if footnote_count > 0:
                self.logger.info(f"  Found {footnote_count} footnotes")
                self.total_footnotes += footnote_count
            
            # Use metadata for title/author or extract from HTML
            title = metadata.get('title') or self._extract_title(html_content, safe_name)
            author = metadata.get('author', 'Unknown')
            
            # Build complete HTML (include MathJax if equations detected)
            complete_html = self._build_html_document(title, author, html_content, has_equations)
            
            # Save HTML
            html_path = article_folder / f"{safe_name}.html"
            html_path.write_text(complete_html, encoding='utf-8')
            
            self.logger.info(f"Saved: {html_path}")
            
        except Exception as e:
            self.logger.error(f"Error converting {docx_path}: {e}", exc_info=True)

    def _convert_with_equation_markers(self, docx_path):
        """Convert document with Office Math equations preserved in place."""
        import tempfile
        import shutil
        import os
        
        # Create a temporary copy
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
            shutil.copy2(docx_path, tmp.name)
            temp_path = tmp.name
        
        equation_map = {}
        
        try:
            # Open with python-docx
            doc = docx.Document(temp_path)
            eq_counter = 0
            
            # Process each paragraph
            for para_idx, paragraph in enumerate(doc.paragraphs):
                # Check if paragraph has Office Math
                math_elements = paragraph._element.xpath('.//m:oMath', namespaces={
                    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math'
                })
                
                if math_elements:
                    # Create a new paragraph text with markers
                    para_text = paragraph.text
                    
                    for math_elem in math_elements:
                        eq_counter += 1
                        marker = f" [EQUATION_{eq_counter}_HERE] "
                        
                        # Extract equation text
                        equation_text = self._extract_math_text_from_element(math_elem)
                        
                        # Store equation for later replacement
                        # Try to make a better LaTeX representation
                        latex_equation = self._convert_to_latex_format(equation_text)
                        equation_map[marker] = latex_equation
                    
                    # Add marker at end of paragraph
                    # (Since we can't easily insert in the middle)
                    if para_text.strip():
                        paragraph.add_run(marker)
                    else:
                        # If paragraph is empty except for equation, just add marker
                        paragraph.text = marker
            
            # Save modified document
            doc.save(temp_path)
            
            # Convert with mammoth
            with open(temp_path, "rb") as docx_file:
                result = mammoth.convert_to_html(
                    docx_file,
                    style_map=self.style_map,
                    convert_image=mammoth.images.img_element(self._image_handler)
                )
            
            html_content = result.value
            
            # Check for warnings (but suppress oMath warnings)
            if result.messages:
                for msg in result.messages:
                    if "oMath" not in str(msg) and "EQUATION_" not in str(msg):
                        self.logger.warning(f"{docx_path.name}: {msg}")
            
            # Replace markers with actual equations
            for marker, equation in equation_map.items():
                html_content = html_content.replace(marker.strip(), equation)
            
            self.logger.info(f"  Processed {len(equation_map)} Office Math equations in place")
            
            return html_content, len(equation_map)
            
        except Exception as e:
            self.logger.error(f"Error in equation marker conversion: {e}")
            # Fall back to regular conversion
            with open(docx_path, "rb") as docx_file:
                result = mammoth.convert_to_html(
                    docx_file,
                    style_map=self.style_map,
                    convert_image=mammoth.images.img_element(self._image_handler)
                )
            return result.value, 0
            
        finally:
            # Clean up temp file
            if os.path.exists(temp_path):
                os.remove(temp_path)

    def _extract_math_text_from_element(self, math_elem):
        """Extract text from a specific math element."""
        text_parts = []
        
        # Find all text nodes within this math element
        for text_elem in math_elem.xpath('.//m:t', namespaces={
            'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math'
        }):
            if text_elem.text:
                text_parts.append(text_elem.text.strip())
        
        return ' '.join(text_parts)

    def _convert_to_latex_format(self, equation_text):
        """Convert extracted equation text to LaTeX format."""
        # This is a basic conversion - you might need to enhance this
        latex = equation_text
        
        # Common replacements
        replacements = [
            ('÷', '\\div'),
            ('×', '\\times'),
            ('±', '\\pm'),
            ('≈', '\\approx'),
            ('≠', '\\neq'),
            ('≤', '\\leq'),
            ('≥', '\\geq'),
            ('∞', '\\infty'),
            ('∑', '\\sum'),
            ('∫', '\\int'),
            ('√', '\\sqrt'),
            ('∂', '\\partial'),
            ('∈', '\\in'),
            ('∉', '\\notin'),
            ('α', '\\alpha'),
            ('β', '\\beta'),
            ('γ', '\\gamma'),
            ('π', '\\pi'),
            ('σ', '\\sigma'),
            ('Σ', '\\Sigma'),
        ]
        
        for old, new in replacements:
            latex = latex.replace(old, new)
        
        # Try to detect fractions
        latex = re.sub(r'(\d+)\s*/\s*(\d+)', r'\\frac{\1}{\2}', latex)
        
        # Return as display equation if it looks complex
        if any(sym in latex for sym in ['\\frac', '\\sum', '\\int', '=']) or len(latex) > 10:
            return f"$${latex}$$"
        else:
            return f"${latex}$"


    def _detect_equation_type(self, docx_path):
        """Detect whether document uses LaTeX or Office Math equations."""
        has_office_math = False
        has_latex = False
        
        try:
            # Check for Office Math
            with zipfile.ZipFile(docx_path, 'r') as zip_file:
                if 'word/document.xml' in zip_file.namelist():
                    with zip_file.open('word/document.xml') as xml_file:
                        content = xml_file.read().decode('utf-8')
                        if 'oMath' in content or 'oMathPara' in content:
                            has_office_math = True
            
            # Check for LaTeX patterns
            with open(docx_path, "rb") as f:
                raw_result = mammoth.extract_raw_text(f)
                raw_text = raw_result.value
                
                latex_patterns = [r'\$[^$\n]+\$', r'\$\$[^$]+\$\$', r'\\[[(]']
                for pattern in latex_patterns:
                    if re.search(pattern, raw_text):
                        has_latex = True
                        break
            
            # Return detected type
            if has_office_math:
                return "office_math"
            elif has_latex:
                return "latex"
            else:
                return "none"
                
        except Exception as e:
            self.logger.debug(f"Error detecting equation type: {e}")
            return "none"

    def _extract_office_math_equations(self, docx_path):
        """Extract Office Math equations from Word document."""
        equations = []
        
        try:
            with zipfile.ZipFile(docx_path, 'r') as zip_file:
                with zip_file.open('word/document.xml') as xml_file:
                    tree = ET.parse(xml_file)
                    root = tree.getroot()
                    
                    # Define namespaces
                    namespaces = {
                        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                        'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math'
                    }
                    
                    # Find all math paragraphs and inline math
                    math_paras = root.findall('.//m:oMathPara', namespaces)
                    inline_maths = root.findall('.//m:oMath', namespaces)
                    
                    eq_counter = 0
                    
                    # Process math paragraphs (display equations)
                    for math_para in math_paras:
                        eq_counter += 1
                        equation_text = self._extract_math_content(math_para, namespaces)
                        equations.append({
                            'id': eq_counter,
                            'type': 'display',
                            'text': equation_text,
                            'latex': f"$${equation_text}$$"
                        })
                    
                    # Process inline math
                    for math in inline_maths:
                        # Skip if already part of a math paragraph
                        parent = math.getparent()
                        while parent is not None:
                            if parent.tag.endswith('oMathPara'):
                                break
                            parent = parent.getparent()
                        
                        if parent is None:  # Not part of oMathPara
                            eq_counter += 1
                            equation_text = self._extract_math_content(math, namespaces)
                            equations.append({
                                'id': eq_counter,
                                'type': 'inline',
                                'text': equation_text,
                                'latex': f"${equation_text}$"
                            })
                    
                    self.logger.info(f"  Extracted {len(equations)} Office Math equations")
                    
        except Exception as e:
            self.logger.warning(f"Could not extract Office Math: {e}")
        
        return equations

    def _extract_math_content(self, math_elem, namespaces):
        """Extract readable content from Office Math XML."""
        text_parts = []
        
        # Extract all text nodes
        for elem in math_elem.iter():
            if elem.text and elem.text.strip():
                text_parts.append(elem.text.strip())
        
        return ' '.join(text_parts)

    def _process_office_math(self, html_content, equations):
        """Add Office Math equations to HTML."""
        if not equations:
            return html_content
        
        # Add equations section at end of document
        equation_html = '\n\n<div class="office-math-equations">\n'
        equation_html += '<h3>معادلات المستند</h3>\n'
        equation_html += '<p class="equation-note">ملاحظة: تم استخراج هذه المعادلات من تنسيق Office Math</p>\n'
        
        for eq in equations:
            if eq['type'] == 'display':
                equation_html += f'<div class="equation display-equation">\n'
                equation_html += f'  {eq["latex"]}\n'
                equation_html += f'</div>\n'
            else:
                equation_html += f'<span class="equation inline-equation">{eq["latex"]}</span> '
        
        equation_html += '</div>\n'
        
        return html_content + equation_html

    def _extract_latex_equations(self, docx_path):
        """Extract LaTeX equations from document text."""
        equations = []
        
        try:
            with open(docx_path, "rb") as f:
                raw_result = mammoth.extract_raw_text(f)
                raw_text = raw_result.value
            
            # Find inline equations
            inline_pattern = r'\$([^$\n]+)\$'
            for match in re.finditer(inline_pattern, raw_text):
                equations.append({
                    'type': 'inline',
                    'latex': match.group(0),
                    'content': match.group(1)
                })
            
            # Find display equations
            display_pattern = r'\$\$([^$]+)\$\$'
            for match in re.finditer(display_pattern, raw_text):
                equations.append({
                    'type': 'display',
                    'latex': match.group(0),
                    'content': match.group(1)
                })
            
            self.logger.info(f"  Extracted {len(equations)} LaTeX equations")
            
        except Exception as e:
            self.logger.debug(f"Error extracting LaTeX: {e}")
        
        return equations
    
    def _extract_metadata(self, docx_path):
        """Extract metadata from document."""
        metadata = {}
        try:
            doc = docx.Document(docx_path)
            if hasattr(doc.core_properties, 'title') and doc.core_properties.title:
                metadata['title'] = doc.core_properties.title
            if hasattr(doc.core_properties, 'author') and doc.core_properties.author:
                metadata['author'] = doc.core_properties.author
            if hasattr(doc.core_properties, 'subject') and doc.core_properties.subject:
                metadata['subject'] = doc.core_properties.subject
        except Exception as e:
            self.logger.debug(f"Could not extract metadata: {e}")
        return metadata
    
    def _image_handler(self, image):
        """Handle image conversion."""
        self.image_counter += 1
        self.total_images += 1
        
        # Get image data
        with image.open() as image_stream:
            image_data = image_stream.read()
        
        # Get image extension
        extension = ".png"
        if hasattr(image, 'content_type'):
            if 'jpeg' in image.content_type:
                extension = ".jpg"
            elif 'png' in image.content_type:
                extension = ".png"
            elif 'gif' in image.content_type:
                extension = ".gif"
        
        # Save image
        filename = f"image_{self.image_counter}{extension}"
        image_path = self.current_image_folder / filename
        
        with open(image_path, "wb") as f:
            f.write(image_data)
        
        # Return with alt text
        return {
            "src": f"images/{filename}",
            "alt": f"صورة {self.image_counter}"
        }
    
    def _preserve_equations(self, html):
        """Preserve and fix LaTeX equations in HTML."""
        # Fix escaped dollar signs
        html = re.sub(r'\\\$', '$', html)
        
        # Fix equations with added spaces
        html = re.sub(r'\$\s+([^$]+?)\s+\$', r'$\1$', html)
        html = re.sub(r'\$\$\s+([^$]+?)\s+\$\$', r'$$\1$$', html)
        
        # Fix HTML entities in equations
        html = html.replace('&lt;', '<')
        html = html.replace('&gt;', '>')
        html = html.replace('&amp;', '&')
        
        # Ensure backslashes are preserved
        html = re.sub(r'\\\\([a-zA-Z])', r'\\\1', html)
        
        return html
    
    def _add_footnote_backlinks(self, html):
        """Add back links from footnotes to text."""
        footnote_pattern = r'<li id="(fn-\d+)">(.*?)</li>'
        
        def add_backlink(match):
            fn_id = match.group(1)
            content = match.group(2)
            fn_number = fn_id.replace('fn-', '')
            return f'<li id="{fn_id}">{content} <a href="#fnref-{fn_number}" class="footnote-backlink" title="العودة إلى النص">↩</a></li>'
        
        return re.sub(footnote_pattern, add_backlink, html, flags=re.DOTALL)
    
    def _enhance_tables(self, html):
        """Add table styling classes."""
        # Add class to tables
        html = html.replace('<table>', '<table class="document-table">')
        # Add responsive wrapper
        html = re.sub(
            r'<table class="document-table">(.*?)</table>',
            r'<div class="table-wrapper"><table class="document-table">\1</table></div>',
            html,
            flags=re.DOTALL
        )
        return html
    
    def _extract_title(self, html, default):
        """Extract title from HTML."""
        match = re.search(r'<h1[^>]*>([^<]+)</h1>', html)
        if match:
            return match.group(1).strip()
        return default
    
    def _build_html_document(self, title, author, body_html, has_equations):
        """Build complete HTML document with all features."""
        # Choose math script based on content
        math_script = ""
        if has_equations:
            math_script = """
    <!-- MathJax for equations -->
    <script>
        window.MathJax = {
            tex: {
                inlineMath: [['$', '$'], ['\\\\(', '\\\\)']],
                displayMath: [['$$', '$$'], ['\\\\[', '\\\\]']],
                processEscapes: true
            },
            svg: {
                fontCache: 'global'
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
    {math_script}
    <style>
        body {{
            font-family: 'Amiri', 'Arial', 'Tahoma', sans-serif;
            line-height: 1.8;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
            direction: rtl;
            text-align: right;
            color: #333;
        }}
        h1, h2, h3, h4 {{
            color: #1a1a1a;
            margin-top: 1.5em;
            margin-bottom: 0.5em;
        }}
        .title {{
            text-align: center;
            font-size: 2.5em;
            margin-bottom: 0.2em;
            color: #0066cc;
        }}
        .subtitle {{
            text-align: center;
            font-size: 1.5em;
            color: #666;
            margin-bottom: 1em;
        }}
        .author {{
            text-align: center;
            color: #666;
            margin-bottom: 2em;
            font-style: italic;
        }}
        img {{
            max-width: 100%;
            height: auto;
            display: block;
            margin: 1em auto;
            border: 1px solid #ddd;
            padding: 5px;
            background: #fff;
        }}
        .caption {{
            text-align: center;
            font-style: italic;
            color: #666;
            font-size: 0.9em;
            margin-top: -0.5em;
            margin-bottom: 1em;
        }}
        .table-wrapper {{
            overflow-x: auto;
            margin: 1em 0;
        }}
        table {{
            border-collapse: collapse;
            width: 100%;
            margin: 1em 0;
            background: #fff;
        }}
        td, th {{
            border: 1px solid #ddd;
            padding: 8px;
            text-align: right;
        }}
        th {{
            background-color: #f5f5f5;
            font-weight: bold;
        }}
        tr:nth-child(even) {{
            background-color: #f9f9f9;
        }}
        blockquote {{
            border-right: 4px solid #ddd;
            margin: 1em 0;
            padding-right: 1em;
            color: #666;
            font-style: italic;
        }}
        /* Footnotes styling */
        .footnotes {{
            margin-top: 3em;
            border-top: 2px solid #ddd;
            padding-top: 1em;
            font-size: 0.9em;
        }}
        sup {{
            font-size: 0.8em;
            color: #0066cc;
        }}
        .footnote-backlink {{
            text-decoration: none;
            margin-right: 0.5em;
            color: #0066cc;
        }}
        /* Equations */
        .office-math-equations {{
            margin-top: 2em;
            padding: 1em;
            background: #f9f9f9;
            border-radius: 5px;
        }}
        .equation {{
            margin: 0.5em 0;
        }}
        .display-equation {{
            text-align: center;
            margin: 1em 0;
        }}
        .MathJax {{
            font-size: 1.1em;
        }}
        /* Print styles */
        @media print {{
            body {{
                margin: 0;
                padding: 10mm;
            }}
            .table-wrapper {{
                overflow: visible;
            }}
        }}
    </style>
</head>
<body>
    <h1 class="title">{title}</h1>
    <p class="author">{author}</p>
    
    <div class="content">
        {body_html}
    </div>
</body>
</html>"""