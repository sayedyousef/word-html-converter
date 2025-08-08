# ============= equation_replacer.py =============
"""
Complete document processor that replaces equations with LaTeX text
Uses pypandoc to maintain document structure while replacing equations
"""

import pypandoc
from pathlib import Path
import re
import hashlib
from typing import Dict, List, Tuple
import tempfile
import json
from logger import setup_logger

logger = setup_logger("equation_replacer")

class DocumentEquationReplacer:
    """
    Replace all equations in Word document with LaTeX plain text + anchors
    Maintains document structure while converting equations
    """
    
    def __init__(self, docx_path: Path):
        self.docx_path = Path(docx_path)
        self.output_path = None
        self.equations_found = []
        self.equation_map = {}
        
    def process_document(self, output_path: Path = None) -> Path:
        """
        Main processing: Convert document replacing equations with LaTeX text
        Returns path to processed document
        """
        if not output_path:
            output_path = self.docx_path.parent / f"{self.docx_path.stem}_latex.docx"
        
        self.output_path = output_path
        
        logger.info(f"Processing {self.docx_path.name}")
        
        # Step 1: Convert to markdown (preserves structure and equations)
        markdown_with_equations = self._convert_to_markdown()
        
        # Step 2: Process equations - replace with anchored LaTeX text
        markdown_processed = self._replace_equations_with_latex(markdown_with_equations)
        
        # Step 3: Convert back to docx
        self._convert_to_docx(markdown_processed, output_path)
        
        # Step 4: Save equation mapping
        self._save_equation_mapping()
        
        logger.info(f"Processed document saved to {output_path.name}")
        logger.info(f"Found and replaced {len(self.equations_found)} equations")
        
        return output_path
    
    def _convert_to_markdown(self) -> str:
        """Convert Word to Markdown preserving equations"""
        logger.info("Converting to markdown...")
        
        # Use pypandoc to convert to markdown with math
        markdown = pypandoc.convert_file(
            str(self.docx_path),
            'markdown',
            format='docx',
            extra_args=[
                '--wrap=preserve',  # Preserve line breaks
                '--extract-media=temp',  # Extract images
                '--standalone',  # Complete document
            ]
        )
        
        return markdown
    
    def _replace_equations_with_latex(self, markdown: str) -> str:
        """
        Find and replace all equations with plain LaTeX text + anchors
        """
        logger.info("Replacing equations with LaTeX text...")
        
        processed = markdown
        equation_counter = 0
        
        # Pattern 1: Display equations $$...$$
        def replace_display_equation(match):
            nonlocal equation_counter
            equation_counter += 1
            
            latex = match.group(1).strip()
            eq_id = self._generate_equation_id(equation_counter, latex)
            
            # Store equation info
            self.equations_found.append({
                'id': eq_id,
                'type': 'display',
                'original': match.group(0),
                'latex': latex,
                'position': equation_counter
            })
            self.equation_map[eq_id] = latex
            
            # Replace with plain text LaTeX + anchor
            # Using HTML comment as anchor that won't show in final doc
            return f'<!-- EQUATION_START id="{eq_id}" type="display" -->\n{latex}\n<!-- EQUATION_END id="{eq_id}" -->'
        
        # Pattern 2: Inline equations $...$
        def replace_inline_equation(match):
            nonlocal equation_counter
            equation_counter += 1
            
            latex = match.group(1).strip()
            eq_id = self._generate_equation_id(equation_counter, latex)
            
            # Store equation info
            self.equations_found.append({
                'id': eq_id,
                'type': 'inline',
                'original': match.group(0),
                'latex': latex,
                'position': equation_counter
            })
            self.equation_map[eq_id] = latex
            
            # Replace with plain text LaTeX + anchor (inline)
            return f'<span class="equation" id="{eq_id}">{latex}</span>'
        
        # Replace display equations
        processed = re.sub(
            r'\$\$(.*?)\$\$',
            replace_display_equation,
            processed,
            flags=re.DOTALL
        )
        
        # Replace inline equations
        processed = re.sub(
            r'\$([^$\n]+)\$',
            replace_inline_equation,
            processed
        )
        
        # Pattern 3: LaTeX environments \begin{equation}...\end{equation}
        def replace_latex_environment(match):
            nonlocal equation_counter
            equation_counter += 1
            
            env_name = match.group(1)
            latex = match.group(2).strip()
            eq_id = self._generate_equation_id(equation_counter, latex)
            
            self.equations_found.append({
                'id': eq_id,
                'type': f'environment_{env_name}',
                'original': match.group(0),
                'latex': latex,
                'position': equation_counter
            })
            self.equation_map[eq_id] = latex
            
            return f'<!-- EQUATION_START id="{eq_id}" type="{env_name}" -->\n{latex}\n<!-- EQUATION_END id="{eq_id}" -->'
        
        # Replace LaTeX environments
        latex_envs = ['equation', 'align', 'gather', 'multline', 'eqnarray']
        for env in latex_envs:
            pattern = rf'\\begin{{{env}}}(.*?)\\end{{{env}}}'
            processed = re.sub(
                pattern,
                lambda m: replace_latex_environment(m) if m else m,
                processed,
                flags=re.DOTALL
            )
        
        return processed
    
    def _generate_equation_id(self, index: int, content: str) -> str:
        """Generate unique ID for equation"""
        # Use content hash for consistency
        hash_val = hashlib.md5(content.encode()).hexdigest()[:6]
        return f"eq_{index}_{hash_val}"
    
    def _convert_to_docx(self, markdown: str, output_path: Path):
        """Convert processed markdown back to Word document"""
        logger.info("Converting back to Word document...")
        
        # Save markdown to temp file
        with tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False, encoding='utf-8') as f:
            f.write(markdown)
            temp_md = f.name
        
        try:
            # Convert markdown to docx
            pypandoc.convert_file(
                temp_md,
                'docx',
                outputfile=str(output_path),
                extra_args=[
                    '--standalone',
                    '--wrap=preserve',
                    '--reference-doc=' + str(self.docx_path),  # Use original as template
                ]
            )
        finally:
            # Clean up temp file
            Path(temp_md).unlink()
    
    def _save_equation_mapping(self):
        """Save equation mapping to JSON file"""
        mapping_file = self.output_path.parent / f"{self.output_path.stem}_equations.json"
        
        data = {
            'source_document': str(self.docx_path),
            'output_document': str(self.output_path),
            'total_equations': len(self.equations_found),
            'equations': self.equations_found
        }
        
        with open(mapping_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        
        logger.info(f"Equation mapping saved to {mapping_file.name}")

