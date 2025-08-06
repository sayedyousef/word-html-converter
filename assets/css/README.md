# CSS Files for Document Processing Project

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
