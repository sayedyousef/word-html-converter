# setup_css_files.py
"""
Complete script to create all CSS files with their full content.
Run this once to set up all CSS files for your project.
"""

from pathlib import Path
import logging

def create_all_css_files():
    """Create all CSS files with complete content."""
    
    css_folder = Path("assets/css")
    css_folder.mkdir(parents=True, exist_ok=True)
    
    print(f"Creating CSS files in {css_folder.absolute()}")
    
    # 1. BASE STYLES
    base_styles = """/* Base styles for all HTML documents */
body {
    font-family: 'Amiri', 'Arial', 'Tahoma', sans-serif;
    line-height: 1.8;
    max-width: 900px;
    margin: 0 auto;
    padding: 20px;
    direction: rtl;
    text-align: right;
    color: #333;
    background-color: #fafafa;
}

h1, h2, h3, h4, h5, h6 {
    color: #1a1a1a;
    margin-top: 1.5em;
    margin-bottom: 0.5em;
    font-weight: bold;
}

h1 {
    font-size: 2.2em;
    border-bottom: 2px solid #e0e0e0;
    padding-bottom: 0.3em;
}

h2 {
    font-size: 1.8em;
    border-bottom: 1px solid #e0e0e0;
    padding-bottom: 0.2em;
}

.title {
    text-align: center;
    font-size: 2.5em;
    margin-bottom: 0.2em;
    color: #0066cc;
    border-bottom: 3px solid #0066cc;
    padding-bottom: 0.5em;
}

.subtitle {
    text-align: center;
    font-size: 1.5em;
    color: #666;
    margin-bottom: 1em;
}

.author {
    text-align: center;
    color: #666;
    margin-bottom: 2em;
    font-style: italic;
}

.content {
    background: white;
    padding: 2em;
    border-radius: 8px;
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
}

p {
    margin: 1em 0;
    text-align: justify;
}

a {
    color: #0066cc;
    text-decoration: none;
}

a:hover {
    text-decoration: underline;
    color: #0052a3;
}

blockquote {
    border-right: 4px solid #ddd;
    margin: 1em 0;
    padding-right: 1em;
    padding-left: 1em;
    color: #666;
    font-style: italic;
    background: #f9f9f9;
    border-radius: 0 5px 5px 0;
}

ul, ol {
    margin: 1em 0;
    padding-right: 2em;
}

li {
    margin: 0.5em 0;
}

pre {
    background: #f5f5f5;
    padding: 1em;
    border-radius: 5px;
    overflow-x: auto;
    direction: ltr;
    text-align: left;
}

code {
    background: #f5f5f5;
    padding: 0.2em 0.4em;
    border-radius: 3px;
    font-family: 'Courier New', Courier, monospace;
    direction: ltr;
}

hr {
    border: none;
    border-top: 2px solid #e0e0e0;
    margin: 2em 0;
}

strong {
    font-weight: bold;
    color: #222;
}

em {
    font-style: italic;
}

::selection {
    background: #b3d4fc;
    text-shadow: none;
}"""
    
    # 2. EQUATION STYLES
    equation_styles = """/* Styles for mathematical equations */
.equation {
    margin: 1em 0;
    position: relative;
}

.display-equation,
.display-math {
    display: block;
    text-align: center;
    margin: 1.5em 0;
    padding: 1em;
    overflow-x: auto;
    background: #f9f9f9;
    border-radius: 5px;
    border: 1px solid #e0e0e0;
}

.inline-equation,
.inline-math {
    display: inline;
    padding: 0 0.3em;
    background: rgba(0, 102, 204, 0.05);
    border-radius: 3px;
}

.office-math-equations {
    margin-top: 2em;
    padding: 1.5em;
    background: #f0f7ff;
    border-radius: 8px;
    border: 1px solid #cce0ff;
}

.office-math-equations h3 {
    color: #0066cc;
    margin-top: 0;
}

.equation-note {
    font-size: 0.9em;
    color: #666;
    font-style: italic;
}

.equation-number {
    position: absolute;
    right: 1em;
    color: #666;
    font-size: 0.9em;
    font-weight: normal;
}

.MathJax {
    font-size: 1.1em;
}

.MathJax_Display {
    margin: 1em 0 !important;
}

.equation-error {
    color: #d00;
    border: 2px solid #d00;
    padding: 0.5em;
    background: #ffe6e6;
    font-family: monospace;
    border-radius: 5px;
}

.equation:hover {
    background: #f0f7ff;
    transition: background 0.3s ease;
}"""
    
    # 3. TABLE STYLES
    table_styles = """/* Table formatting and responsive tables */
.table-wrapper {
    overflow-x: auto;
    margin: 1.5em 0;
    border-radius: 8px;
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
}

table {
    border-collapse: collapse;
    width: 100%;
    margin: 1em 0;
    background: white;
    font-size: 0.95em;
}

.document-table {
    border: 1px solid #ddd;
}

td, th {
    border: 1px solid #ddd;
    padding: 10px 12px;
    text-align: right;
    vertical-align: middle;
}

th {
    background-color: #f5f5f5;
    font-weight: bold;
    color: #333;
    background: linear-gradient(to bottom, #f9f9f9, #e9e9e9);
}

thead {
    background: #f0f0f0;
}

thead th {
    position: sticky;
    top: 0;
    z-index: 10;
    box-shadow: 0 2px 2px -1px rgba(0, 0, 0, 0.1);
}

tbody tr:nth-child(even) {
    background-color: #f9f9f9;
}

tbody tr:nth-child(odd) {
    background-color: white;
}

tbody tr:hover {
    background-color: #e6f3ff;
    transition: background-color 0.2s ease;
}

caption {
    padding: 0.5em;
    color: #666;
    font-style: italic;
    caption-side: bottom;
}

.table-compact td,
.table-compact th {
    padding: 5px 8px;
}

.table-bordered {
    border: 2px solid #333;
}

.table-bordered td,
.table-bordered th {
    border: 1px solid #333;
}"""
    
    # 4. IMAGE STYLES
    image_styles = """/* Image display and caption styles */
img {
    max-width: 100%;
    height: auto;
    display: block;
    margin: 1.5em auto;
    border: 1px solid #ddd;
    padding: 5px;
    background: white;
    border-radius: 5px;
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
}

img:hover {
    box-shadow: 0 4px 8px rgba(0, 0, 0, 0.15);
    transition: box-shadow 0.3s ease;
}

figure {
    margin: 2em 0;
    text-align: center;
}

.caption,
figcaption {
    text-align: center;
    font-style: italic;
    color: #666;
    font-size: 0.9em;
    margin-top: 0.5em;
    margin-bottom: 1em;
    padding: 0 1em;
}

.image-gallery {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
    gap: 1em;
    margin: 2em 0;
}

.image-gallery img {
    width: 100%;
    margin: 0;
}

.inline-image {
    display: inline-block;
    margin: 0 0.5em;
    vertical-align: middle;
    max-height: 2em;
}

.large-image {
    max-width: none;
    width: 100%;
}

img[data-anchor] {
    scroll-margin-top: 20px;
}

img[data-anchor]:target {
    border: 3px solid #0066cc !important;
    box-shadow: 0 0 15px rgba(0, 102, 204, 0.5);
    animation: highlight 1s ease;
}

@keyframes highlight {
    0% { transform: scale(1); }
    50% { transform: scale(1.02); }
    100% { transform: scale(1); }
}"""
    
    # 5. FOOTNOTE STYLES
    footnote_styles = """/* Footnote and endnote formatting */
.footnotes {
    margin-top: 3em;
    border-top: 2px solid #ddd;
    padding-top: 1em;
    font-size: 0.9em;
}

sup {
    font-size: 0.75em;
    color: #0066cc;
    font-weight: bold;
    padding: 0 0.1em;
}

sup a {
    color: #0066cc;
    text-decoration: none;
}

sup a:hover {
    text-decoration: underline;
}

.footnotes ol {
    padding-right: 1.5em;
}

.footnotes li {
    margin: 0.5em 0;
    line-height: 1.6;
}

.footnote-backlink {
    text-decoration: none;
    margin-right: 0.5em;
    color: #0066cc;
    font-size: 1.1em;
}

.footnote-backlink:hover {
    text-decoration: none;
    color: #0052a3;
}

.footnotes li:target {
    background: #fffbcc;
    padding: 0.5em;
    border-radius: 5px;
    animation: footnote-highlight 2s ease;
}

@keyframes footnote-highlight {
    0% { background: #ffeb3b; }
    100% { background: #fffbcc; }
}

.endnotes {
    margin-top: 4em;
    padding-top: 2em;
    border-top: 3px double #999;
}

.endnotes h2 {
    color: #666;
    font-size: 1.3em;
}"""
    
    # 6. ANCHOR STYLES
    anchor_styles = """/* Anchor and navigation styles */
.equation-anchor {
    display: inline-block;
    width: 0;
    height: 0;
    visibility: hidden;
    scroll-margin-top: 30px;
}

.equation-anchor:target {
    background: yellow;
    padding: 5px;
    visibility: visible;
    width: auto;
    height: auto;
    animation: anchor-pulse 2s ease;
}

@keyframes anchor-pulse {
    0% { background: #ffeb3b; }
    50% { background: #fff59d; }
    100% { background: yellow; }
}

.nav-anchor {
    scroll-margin-top: 50px;
}

.toc-anchor {
    scroll-margin-top: 20px;
}

html {
    scroll-behavior: smooth;
}

a.anchor-link {
    color: #999;
    font-size: 0.8em;
    margin-left: 0.5em;
    opacity: 0;
    transition: opacity 0.3s ease;
}

h1:hover a.anchor-link,
h2:hover a.anchor-link,
h3:hover a.anchor-link {
    opacity: 1;
}

.skip-link {
    position: absolute;
    top: -40px;
    left: 0;
    background: #0066cc;
    color: white;
    padding: 8px;
    text-decoration: none;
    z-index: 100;
}

.skip-link:focus {
    top: 0;
}

.bookmark-indicator {
    position: absolute;
    left: -25px;
    color: #0066cc;
    font-size: 1.2em;
}"""
    
    # 7. PRINT STYLES
    print_styles = """/* Print-specific styles */
@media print {
    @page {
        margin: 2cm;
        size: A4;
    }
    
    body {
        margin: 0;
        padding: 0;
        font-size: 11pt;
        line-height: 1.5;
        color: black;
        background: white;
    }
    
    .content {
        box-shadow: none;
        padding: 0;
    }
    
    .no-print,
    .equation-anchor,
    .footnote-backlink,
    .skip-link,
    nav,
    .navigation {
        display: none !important;
    }
    
    h1, h2, h3, h4, h5, h6 {
        page-break-after: avoid;
        page-break-inside: avoid;
    }
    
    h1 { font-size: 18pt; }
    h2 { font-size: 16pt; }
    h3 { font-size: 14pt; }
    
    p, blockquote, pre {
        orphans: 3;
        widows: 3;
    }
    
    .table-wrapper {
        overflow: visible !important;
    }
    
    table {
        page-break-inside: avoid;
    }
    
    thead {
        display: table-header-group;
    }
    
    img {
        max-width: 100% !important;
        page-break-inside: avoid;
        border: 1px solid #999;
    }
    
    .display-equation,
    .display-math {
        page-break-inside: avoid;
        background: white;
        border: 1px solid #ccc;
    }
    
    a {
        color: black;
        text-decoration: underline;
    }
    
    a[href^="http"]:after {
        content: " (" attr(href) ")";
        font-size: 0.8em;
        color: #666;
    }
    
    .footnotes {
        page-break-before: always;
    }
    
    pre {
        white-space: pre-wrap;
        word-wrap: break-word;
    }
}"""
    
    # 8. RESPONSIVE STYLES
    responsive_styles = """/* Mobile and tablet responsive design */
@media screen and (max-width: 768px) {
    body {
        padding: 15px;
        font-size: 15px;
    }
    
    .content {
        padding: 1em;
    }
    
    h1 { font-size: 1.8em; }
    h2 { font-size: 1.5em; }
    
    .title {
        font-size: 2em;
    }
    
    table {
        font-size: 0.9em;
    }
    
    td, th {
        padding: 8px;
    }
    
    .display-equation,
    .display-math {
        padding: 0.5em;
        font-size: 0.95em;
    }
}

@media screen and (max-width: 480px) {
    body {
        padding: 10px;
        font-size: 14px;
        max-width: 100%;
    }
    
    .content {
        padding: 0.5em;
        border-radius: 0;
    }
    
    h1 { font-size: 1.5em; }
    h2 { font-size: 1.3em; }
    
    .title {
        font-size: 1.6em;
    }
    
    .table-wrapper {
        margin: 1em -10px;
        padding: 0 10px;
    }
    
    table {
        font-size: 0.85em;
        min-width: 100%;
    }
    
    td, th {
        padding: 5px;
    }
    
    .display-equation,
    .display-math {
        padding: 0.3em;
        font-size: 0.9em;
        margin: 1em 0;
    }
    
    img {
        padding: 2px;
        margin: 1em 0;
    }
    
    .footnotes {
        font-size: 0.85em;
    }
    
    pre {
        padding: 0.5em;
        font-size: 0.85em;
    }
}

@media screen and (max-width: 640px) and (orientation: landscape) {
    body {
        padding: 10px 20px;
    }
}

@media (-webkit-min-device-pixel-ratio: 2), (min-resolution: 192dpi) {
    img {
        image-rendering: -webkit-optimize-contrast;
        image-rendering: crisp-edges;
    }
}"""
    
    # 9. THEME STYLES
    theme_styles = """/* Optional theme variations */
@media (prefers-color-scheme: dark) {
    body.auto-dark {
        background: #1a1a1a;
        color: #e0e0e0;
    }
    
    body.auto-dark .content {
        background: #2a2a2a;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.5);
    }
    
    body.auto-dark h1,
    body.auto-dark h2,
    body.auto-dark h3,
    body.auto-dark h4 {
        color: #f0f0f0;
    }
    
    body.auto-dark .title {
        color: #4da6ff;
    }
    
    body.auto-dark table {
        background: #2a2a2a;
    }
    
    body.auto-dark th {
        background: #333;
        color: #f0f0f0;
    }
    
    body.auto-dark td {
        border-color: #444;
    }
    
    body.auto-dark .display-equation,
    body.auto-dark .display-math {
        background: #333;
        border-color: #444;
        color: #e0e0e0;
    }
    
    body.auto-dark blockquote {
        background: #2a2a2a;
        border-color: #666;
    }
    
    body.auto-dark code,
    body.auto-dark pre {
        background: #333;
        color: #f0f0f0;
    }
}

body.high-contrast {
    background: white;
    color: black;
    font-weight: 500;
}

body.high-contrast a {
    color: #0000ff;
    text-decoration: underline;
}

body.high-contrast .content {
    border: 2px solid black;
}

body.high-contrast table,
body.high-contrast th,
body.high-contrast td {
    border: 2px solid black;
}

body.large-text {
    font-size: 20px;
    line-height: 2;
}

body.large-text h1 {
    font-size: 2.5em;
}

body.large-text h2 {
    font-size: 2em;
}"""
    
    # 10. UTILITIES
    utilities = """/* Utility classes */
.text-left { text-align: left !important; }
.text-right { text-align: right !important; }
.text-center { text-align: center !important; }
.text-justify { text-align: justify !important; }

.d-none { display: none !important; }
.d-block { display: block !important; }
.d-inline { display: inline !important; }
.d-inline-block { display: inline-block !important; }

.m-0 { margin: 0 !important; }
.m-1 { margin: 0.5em !important; }
.m-2 { margin: 1em !important; }
.m-3 { margin: 1.5em !important; }
.m-4 { margin: 2em !important; }

.mt-0 { margin-top: 0 !important; }
.mt-1 { margin-top: 0.5em !important; }
.mt-2 { margin-top: 1em !important; }
.mt-3 { margin-top: 1.5em !important; }
.mt-4 { margin-top: 2em !important; }

.mb-0 { margin-bottom: 0 !important; }
.mb-1 { margin-bottom: 0.5em !important; }
.mb-2 { margin-bottom: 1em !important; }
.mb-3 { margin-bottom: 1.5em !important; }
.mb-4 { margin-bottom: 2em !important; }

.p-0 { padding: 0 !important; }
.p-1 { padding: 0.5em !important; }
.p-2 { padding: 1em !important; }
.p-3 { padding: 1.5em !important; }
.p-4 { padding: 2em !important; }

.text-primary { color: #0066cc !important; }
.text-secondary { color: #666 !important; }
.text-success { color: #28a745 !important; }
.text-danger { color: #dc3545 !important; }
.text-warning { color: #ffc107 !important; }
.text-info { color: #17a2b8 !important; }
.text-muted { color: #999 !important; }

.bg-primary { background-color: #0066cc !important; }
.bg-secondary { background-color: #666 !important; }
.bg-success { background-color: #28a745 !important; }
.bg-danger { background-color: #dc3545 !important; }
.bg-warning { background-color: #ffc107 !important; }
.bg-info { background-color: #17a2b8 !important; }
.bg-light { background-color: #f8f9fa !important; }
.bg-dark { background-color: #343a40 !important; }

.border { border: 1px solid #dee2e6 !important; }
.border-0 { border: 0 !important; }
.border-top { border-top: 1px solid #dee2e6 !important; }
.border-bottom { border-bottom: 1px solid #dee2e6 !important; }
.border-left { border-left: 1px solid #dee2e6 !important; }
.border-right { border-right: 1px solid #dee2e6 !important; }

.rounded { border-radius: 5px !important; }
.rounded-0 { border-radius: 0 !important; }
.rounded-circle { border-radius: 50% !important; }

.float-left { float: left !important; }
.float-right { float: right !important; }
.float-none { float: none !important; }

.clearfix::after {
    display: block;
    clear: both;
    content: "";
}

.w-25 { width: 25% !important; }
.w-50 { width: 50% !important; }
.w-75 { width: 75% !important; }
.w-100 { width: 100% !important; }
.w-auto { width: auto !important; }

.h-25 { height: 25% !important; }
.h-50 { height: 50% !important; }
.h-75 { height: 75% !important; }
.h-100 { height: 100% !important; }
.h-auto { height: auto !important; }

.visible { visibility: visible !important; }
.invisible { visibility: hidden !important; }

.opacity-0 { opacity: 0 !important; }
.opacity-25 { opacity: 0.25 !important; }
.opacity-50 { opacity: 0.5 !important; }
.opacity-75 { opacity: 0.75 !important; }
.opacity-100 { opacity: 1 !important; }

.shadow-none { box-shadow: none !important; }
.shadow-sm { box-shadow: 0 1px 2px rgba(0,0,0,0.075) !important; }
.shadow { box-shadow: 0 2px 4px rgba(0,0,0,0.1) !important; }
.shadow-lg { box-shadow: 0 4px 8px rgba(0,0,0,0.15) !important; }

.sr-only {
    position: absolute;
    width: 1px;
    height: 1px;
    padding: 0;
    margin: -1px;
    overflow: hidden;
    clip: rect(0,0,0,0);
    white-space: nowrap;
    border: 0;
}"""
    
    # Create all CSS files
    css_files = {
        'base-styles.css': base_styles,
        'equation-styles.css': equation_styles,
        'table-styles.css': table_styles,
        'image-styles.css': image_styles,
        'footnote-styles.css': footnote_styles,
        'anchor-styles.css': anchor_styles,
        'print-styles.css': print_styles,
        'responsive-styles.css': responsive_styles,
        'theme-styles.css': theme_styles,
        'utilities.css': utilities
    }
    
    created_files = []
    for filename, content in css_files.items():
        filepath = css_folder / filename
        filepath.write_text(content, encoding='utf-8')
        created_files.append(filepath)
        print(f"✅ Created: {filepath}")
    
    # Create a master CSS file that imports all others
    master_css = """/* Master CSS file - imports all other stylesheets */
@import url('base-styles.css');
@import url('equation-styles.css');
@import url('table-styles.css');
@import url('image-styles.css');
@import url('footnote-styles.css');
@import url('anchor-styles.css');
@import url('responsive-styles.css');
@import url('print-styles.css');
@import url('theme-styles.css');
@import url('utilities.css');

/* Custom overrides can be added here */
"""
    
    master_path = css_folder / 'master.css'
    master_path.write_text(master_css, encoding='utf-8')
    print(f"✅ Created: {master_path}")
    
    print(f"\n🎉 Successfully created {len(css_files) + 1} CSS files in {css_folder.absolute()}")
    print("\n📁 File structure:")
    print(f"   {css_folder}/")
    for file in created_files:
        print(f"   ├── {file.name}")
    print(f"   └── master.css")
    
    # Create a README for the CSS folder
    readme_content = """# CSS Files for Document Processing Project

## File Descriptions

- **base-styles.css**: Core typography, layout, and document structure
- **equation-styles.css**: Mathematical equation formatting (LaTeX, MathJax)
- **table-styles.css**: Table formatting and responsive tables
- **image-styles.css**: Image display, captions, and galleries
- **footnote-styles.css**: Footnote and endnote formatting
- **anchor-styles.css**: Anchor links and navigation
- **print-styles.css**: Print-specific optimizations
- **responsive-styles.css**: Mobile and tablet responsive design
- **theme-styles.css**: Dark mode and accessibility themes
- **utilities.css**: Utility classes for quick styling
- **master.css**: Imports all other CSS files

## Usage

### Option 1: Link individual files
```html
<link rel="stylesheet" href="assets/css/base-styles.css">
<link rel="stylesheet" href="assets/css/equation-styles.css">
<!-- Add other files as needed -->
```

### Option 2: Use master file
```html
<link rel="stylesheet" href="assets/css/master.css">
```

## Customization

- Modify individual CSS files for specific components
- Add custom overrides at the end of master.css
- Use utility classes for quick adjustments

## RTL Support

All styles are optimized for Right-to-Left (RTL) Arabic text by default.
"""
    
    readme_path = css_folder / 'README.md'
    readme_path.write_text(readme_content, encoding='utf-8')
    print(f"\n📖 Created README: {readme_path}")
    
    return css_folder

if __name__ == "__main__":
    css_folder = create_all_css_files()
    print(f"\n✨ Setup complete! CSS files are ready to use.")
    print(f"📍 Location: {css_folder.absolute()}")
