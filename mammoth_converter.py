# mammoth_converter.py
"""Enhanced Word to HTML converter using mammoth with equation fixes and anchor support."""

import mammoth
import logging
import re
import docx
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from utils import sanitize_filename, format_article_number, detect_latex_equations
import tempfile
import shutil
import os
import json
from config import Config

class MammothConverter:
    """Enhanced converter using mammoth with all features."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.total_equations = 0
        self.total_footnotes = 0
        self.total_images = 0
        self.input_folder = None
        self.output_folder = None
        
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
        
        # Initialize anchor registry
        self.anchor_registry = {}
        self.equation_positions = {}
    
    def convert_folder(self, input_folder: Path, output_folder: Path):
        """Convert all Word documents in folder."""
        self.input_folder = input_folder
        self.output_folder = output_folder  

        # Find all .docx files
        docx_files = list(input_folder.rglob("*.docx"))
        docx_files = [f for f in docx_files if not f.name.startswith("~")]
        
        self.logger.info(f"Found {len(docx_files)} documents")
        
        # Track statistics
        #self.total_equations = 0
        #self.total_footnotes = 0
        #self.total_images = 0
        
        for idx, docx_file in enumerate(docx_files, 1):
            self._convert_document(docx_file, output_folder, idx)
        
        # Log summary
        self.logger.info("=" * 60)
        self.logger.info("Conversion Summary:")
        self.logger.info(f"  Total equations found: {self.total_equations}")
        self.logger.info(f"  Total footnotes: {self.total_footnotes}")
        self.logger.info(f"  Total images: {self.total_images}")
        self.logger.info("=" * 60)
    
    def _convert_document(self, docx_path: Path, output_base: Path, index: int, 
                     input_folder: Path = None, output_folder: Path = None):
        """Convert single document with enhanced features."""
        try:
            self.logger.info(f"Converting [{index}]: {docx_path.name}")
            
            # Initialize anchor registry for this document
            self.anchor_registry = {}
            
            # Extract metadata using python-docx
            metadata = self._extract_metadata(docx_path)
            
            # Get relative path to preserve folder structure
            #self.input_folder = input_folder or self.input_folder
            #self.input_folder = Path(input_folder) if input_folder else self.input_folder

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
            
            # Handle Office Math equations with ENHANCED position preservation
            if equation_type == "office_math":
                html_content, equation_count = self._convert_with_equation_markers_fixed(docx_path)
                self.total_equations += equation_count
                has_equations = equation_count > 0
            else:
                # Regular conversion for LaTeX or no equations
                with open(docx_path, "rb") as docx_file:
                    result = mammoth.convert_to_html(
                        docx_file,
                        style_map=self.style_map,
                        convert_image=mammoth.images.img_element(self._image_handler_with_anchor)
                    )
                
                html_content = result.value
                
                # Check for warnings (but suppress oMath warnings)
                if result.messages:
                    for msg in result.messages:
                        if "oMath" not in str(msg):
                            self.logger.warning(f"{docx_path.name}: {msg}")
                
                # Process LaTeX equations if present
                if equation_type == "latex":
                    html_content = self._preserve_equations_with_anchors(html_content)
                
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
            
            # Build complete HTML with ENHANCED features
            complete_html = self._build_html_document_enhanced(title, author, html_content, has_equations)
            
            # Save HTML
            html_path = article_folder / f"{safe_name}.html"
            html_path.write_text(complete_html, encoding='utf-8')
            
            # Save anchor registry as JSON
            if self.anchor_registry:
                anchor_path = article_folder / f"{safe_name}.anchors.json"
                with open(anchor_path, 'w', encoding='utf-8') as f:
                    json.dump(self.anchor_registry, f, indent=2, ensure_ascii=False)
                self.logger.info(f"  Saved {len(self.anchor_registry)} anchors")
            
            self.logger.info(f"Saved: {html_path}")
            
        except Exception as e:
            self.logger.error(f"Error converting {docx_path}: {e}", exc_info=True)

    def _convert_with_equation_markers_fixed(self, docx_path):
        """Convert document with Office Math equations - FIXED VERSION."""
        import tempfile
        import os
        import shutil
        from lxml import etree
        import zipfile
        
        self.logger.info("Converting with equation markers (fixed)...")
        
        # Create temp copy
        temp_fd, temp_path = tempfile.mkstemp(suffix='.docx')
        os.close(temp_fd)
        shutil.copy2(docx_path, temp_path)
        
        equation_map = {}
        equation_positions = {}
        equation_counter = 0
        
        try:
            # Process the document XML
            with zipfile.ZipFile(temp_path, 'r') as zip_ref:
                # Read document.xml
                doc_xml = zip_ref.read('word/document.xml')
                root = etree.fromstring(doc_xml)
                
                # Define namespaces
                ns = {
                    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math'
                }
                
                # Find all paragraphs with math
                para_index = 0
                for para in root.findall('.//w:p', ns):
                    para_index += 1
                    
                    # Check for Office Math elements
                    math_elements = para.findall('.//m:oMath', ns)
                    
                    if math_elements:
                        for math_idx, math_elem in enumerate(math_elements):
                            equation_counter += 1
                            
                            # Extract text from math element using findall instead of xpath
                            text_parts = []
                            for t_elem in math_elem.findall('.//m:t', ns):
                                if t_elem.text:
                                    text_parts.append(t_elem.text.strip())
                            
                            equation_text = ' '.join(text_parts)
                            
                            # Convert to LaTeX
                            latex_equation = self._convert_to_latex_format_enhanced(equation_text)
                            
                            # Create unique marker
                            marker = f"[EQUATION_{equation_counter}]"
                            anchor_id = f"eq-{equation_counter}"
                            
                            # Store in map
                            equation_map[marker] = {
                                'latex': latex_equation,
                                'anchor': anchor_id,
                                'original': equation_text
                            }
                            
                            # Track position
                            equation_positions[equation_counter] = {
                                'paragraph': para_index,
                                'index_in_para': math_idx
                            }
                            
                            # Replace math element with marker text
                            # Create text element with marker
                            marker_elem = etree.Element('{%s}t' % ns['w'])
                            marker_elem.text = marker
                            
                            # Replace math element with text run containing marker
                            run = etree.Element('{%s}r' % ns['w'])
                            run.append(marker_elem)
                            
                            # Replace the math element
                            parent = math_elem.getparent()
                            parent.replace(math_elem, run)
                
                # Write modified XML back
                modified_xml = etree.tostring(root, encoding='unicode')
                
                # Update the docx file
                with zipfile.ZipFile(temp_path, 'a') as zip_out:
                    # Remove old document.xml
                    zip_out.writestr('word/document.xml', modified_xml)
            
            # Now convert with mammoth
            with open(temp_path, "rb") as docx_file:
                result = mammoth.convert_to_html(
                    docx_file,
                    style_map=self.style_map,
                    convert_image=mammoth.images.img_element(self._image_handler_with_anchor)
                )
                html_content = result.value
                
                # Log conversion messages
                for msg in result.messages:
                    if msg.type == 'warning':
                        self.logger.warning(f"{docx_path.name}: {msg}")
            
            # Replace markers with equations AND anchors
            for marker, eq_data in equation_map.items():
                anchor_html = f'<a id="{eq_data["anchor"]}" class="equation-anchor"></a>'
                
                # Wrap equation with proper tags
                if '$$' in eq_data['latex'] or len(eq_data['latex']) > 50:
                    equation_html = f'{anchor_html}<div class="equation display-math">$${eq_data["latex"]}$$</div>'
                else:
                    equation_html = f'{anchor_html}<span class="equation inline-math">${eq_data["latex"]}$</span>'
                
                # Replace marker in HTML
                html_content = html_content.replace(marker, equation_html)
            
            self.logger.info(f"  Processed {len(equation_map)} Office Math equations with anchors")
            
            # Store positions for later reference
            self.equation_positions = equation_positions
            
            return html_content, len(equation_map)
            
        except Exception as e:
            self.logger.error(f"Error in equation marker conversion: {e}")
            # Fall back to regular conversion without equations
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


    def create_word_documents_with_anchors(self):
        """Create Word documents identical to originals but with anchors added."""
        
        self.logger.info("=" * 60)
        self.logger.info("Creating Word Documents with Anchors")
        self.logger.info("=" * 60)
        from word_anchor_adder import WordAnchorAdder
        
        anchor_adder = WordAnchorAdder()
        anchor_adder.process_folder(self.input_folder, self.output_folder)
        
        self.logger.info("Finished creating anchored Word documents")




    def _extract_math_text_from_element(self, math_elem):
        """Extract text from a specific math element - FIXED VERSION."""
        text_parts = []
        
        # Instead of using xpath with namespaces, use a different approach
        # Option 1: Use find with namespaces
        ns = {'m': 'http://schemas.openxmlformats.org/officeDocument/2006/math'}
        
        # Find all text elements within this math element
        for text_elem in math_elem.findall('.//m:t', ns):
            if text_elem.text:
                text_parts.append(text_elem.text.strip())
        
        return ' '.join(text_parts)


    def _convert_to_latex_format_enhanced(self, equation_text):
        """Enhanced version - Convert extracted equation text to LaTeX format."""
        # Extended replacements for more symbols
        replacements = [
            ('÷', '\\div'), ('×', '\\times'), ('±', '\\pm'),
            ('≈', '\\approx'), ('≠', '\\neq'), ('≤', '\\leq'),
            ('≥', '\\geq'), ('∞', '\\infty'), ('∑', '\\sum'),
            ('∫', '\\int'), ('√', '\\sqrt'), ('∂', '\\partial'),
            ('∈', '\\in'), ('∉', '\\notin'), ('∅', '\\emptyset'),
            ('α', '\\alpha'), ('β', '\\beta'), ('γ', '\\gamma'),
            ('δ', '\\delta'), ('ε', '\\epsilon'), ('θ', '\\theta'),
            ('λ', '\\lambda'), ('μ', '\\mu'), ('π', '\\pi'),
            ('σ', '\\sigma'), ('τ', '\\tau'), ('φ', '\\phi'),
            ('ω', '\\omega'), ('Σ', '\\Sigma'), ('Δ', '\\Delta'),
            ('Ω', '\\Omega'), ('→', '\\rightarrow'), ('←', '\\leftarrow'),
            ('⇒', '\\Rightarrow'), ('⇔', '\\Leftrightarrow'),
        ]
        
        latex = equation_text
        for old, new in replacements:
            latex = latex.replace(old, new)
        
        # Detect and convert fractions
        latex = re.sub(r'(\d+)\s*/\s*(\d+)', r'\\frac{\1}{\2}', latex)
        
        # Detect exponents and subscripts
        latex = re.sub(r'(\w+)\^(\d+)', r'\1^{\2}', latex)
        latex = re.sub(r'(\w+)_(\d+)', r'\1_{\2}', latex)
        
        # Return as display equation if it looks complex
        if any(sym in latex for sym in ['\\frac', '\\sum', '\\int', '=']) or len(latex) > 15:
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
                
                #latex_patterns = [r'\$[^$\n]+\$', r'\$\$[^$]+\$\$', r'\\[[(]']
                latex_patterns = [r'\$[^$\n]+\$', r'\$\$[^$]+\$\$', r'\\\[', r'\\\(']

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
        """Handle image conversion (fallback without anchors)."""
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
    
    def _image_handler_with_anchor(self, image):
        """Enhanced image handler that adds anchors."""
        self.image_counter += 1
        self.total_images += 1
        
        # Generate anchor ID
        anchor_id = f"img-anchor-{self.image_counter}"
        
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
        
        # Register anchor
        self.anchor_registry[anchor_id] = {
            'type': 'image',
            'filename': filename,
            'path': str(image_path)
        }
        
        # Return with anchor data attribute
        return {
            "src": f"images/{filename}",
            "alt": f"صورة {self.image_counter}",
            "data-anchor": anchor_id  # Add anchor as data attribute
        }
    
    def _preserve_equations(self, html):
        """Preserve and fix LaTeX equations in HTML (without anchors)."""
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
    
    def _preserve_equations_with_anchors(self, html):
        """Preserve LaTeX equations and add anchors."""
        equation_counter = 0
        
        def add_anchor_to_display_equation(match):
            nonlocal equation_counter
            equation_counter += 1
            anchor_id = f"latex-anchor-{equation_counter}"
            anchor_html = f'<a id="{anchor_id}" class="equation-anchor"></a>'
            
            # Register anchor
            self.anchor_registry[anchor_id] = {
                'type': 'latex_equation',
                'format': 'display',
                'content': match.group(0)
            }
            
            return f'{anchor_html}<div class="equation display-math">{match.group(0)}</div>'
        
        def add_anchor_to_inline_equation(match):
            nonlocal equation_counter
            equation_counter += 1
            anchor_id = f"latex-inline-{equation_counter}"
            anchor_html = f'<a id="{anchor_id}" class="equation-anchor"></a>'
            
            # Register anchor
            self.anchor_registry[anchor_id] = {
                'type': 'latex_equation',
                'format': 'inline',
                'content': match.group(0)
            }
            
            return f'{anchor_html}<span class="equation inline-math">{match.group(0)}</span>'
        
        # Add anchors to display equations
        html = re.sub(r'(\$\$[^$]+\$\$)', add_anchor_to_display_equation, html)
        
        # Add anchors to inline equations
        html = re.sub(r'(\$[^$\n]+\$)', add_anchor_to_inline_equation, html)
        
        # Fix other issues
        html = re.sub(r'\\\$', '$', html)
        html = re.sub(r'\$\s+([^$]+?)\s+\$', r'$\1$', html)
        html = html.replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&')
        html = re.sub(r'\\\\([a-zA-Z])', r'\\\1', html)
        
        return html
    
    def _add_footnote_backlinks(self, html):
        """Add back links from footnotes to text."""
        footnote_pattern = r'<li id="(fn-\d+)">(.*?)</li>'
        
        def add_backlink(match):
            fn_id = match.group(1)
            content = match.group(2)
            fn_number = fn_id.replace('fn-', '')
            
            # Register footnote anchor
            anchor_id = f"footnote-{fn_number}"
            self.anchor_registry[anchor_id] = {
                'type': 'footnote',
                'number': fn_number,
                'id': fn_id
            }
            
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
        """Build complete HTML document with all features (original version)."""
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

    def _build_html_document_enhanced(self, title, author, body_html, has_equations):
        """Enhanced HTML builder with anchor support."""
        # MathJax script with better configuration
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
""" if has_equations else ""
        
        # Enhanced CSS with anchor styles
        anchor_styles = """
        /* Anchor styles */
        .equation-anchor {
            display: inline-block;
            width: 0;
            height: 0;
            visibility: hidden;
        }
        
        .equation-anchor:target {
            background: yellow;
            padding: 5px;
            visibility: visible;
            width: auto;
            height: auto;
        }
        
        img[data-anchor] {
            scroll-margin-top: 20px;
        }
        
        img[data-anchor]:target {
            border: 3px solid #0066cc !important;
            box-shadow: 0 0 10px rgba(0, 102, 204, 0.5);
        }
        
        /* Enhanced equation styles */
        .equation {
            position: relative;
            margin: 0.5em 0;
        }
        
        .display-math {
            display: block;
            text-align: center;
            margin: 1em 0;
            padding: 0.5em;
            overflow-x: auto;
        }
        
        .inline-math {
            display: inline;
            padding: 0 0.2em;
        }
        
        /* Equation numbering */
        .equation-number {
            position: absolute;
            right: 0;
            color: #666;
            font-size: 0.9em;
        }
        
        /* Error handling for equations */
        .equation-error {
            color: red;
            border: 1px solid red;
            padding: 0.5em;
            background: #ffe6e6;
            font-family: monospace;
        }
"""
        
        # Get base HTML from original method
        base_html = self._build_html_document(title, author, body_html, has_equations)
        
        # Insert enhanced styles before closing </style> tag
        enhanced_html = base_html.replace('</style>', f'{anchor_styles}\n    </style>')
        
        # Add JavaScript for equation numbering and error handling
        js_enhancements = """
    <!-- Enhanced JavaScript for equations and anchors -->
    <script>
        // Handle MathJax errors gracefully
        document.addEventListener('DOMContentLoaded', function() {
            if (window.MathJax) {
                window.MathJax.startup.promise.catch(function (e) {
                    console.error('MathJax startup failed:', e);
                });
            }
            
            // Add equation numbering
            const displayEquations = document.querySelectorAll('.display-math');
            displayEquations.forEach((eq, index) => {
                if (!eq.querySelector('.equation-number')) {
                    const number = document.createElement('span');
                    number.className = 'equation-number';
                    number.textContent = `(${index + 1})`;
                    eq.appendChild(number);
                }
            });
            
            // Smooth scroll to anchors
            if (window.location.hash) {
                const target = document.querySelector(window.location.hash);
                if (target) {
                    setTimeout(() => {
                        target.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    }, 500);
                }
            }
        });
    </script>
"""
        
        # Insert JavaScript before closing </body> tag
        enhanced_html = enhanced_html.replace('</body>', f'{js_enhancements}\n</body>')
        
        return enhanced_html