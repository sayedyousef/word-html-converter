# models.py
"""Data models for document conversion."""

from dataclasses import dataclass, field
from typing import List, Dict, Optional
from pathlib import Path

@dataclass
class ImageInfo:
    """Information about an image in document."""
    original_path: Optional[Path] = None
    new_filename: str = ""
    caption: str = ""
    alt_text: str = ""
    number: int = 0

@dataclass
class FootnoteInfo:
    """Information about a footnote."""
    id: str
    text: str
    contains_latex: bool = False

@dataclass
class DocumentContent:
    """Represents parsed document content."""
    title: str
    author: str
    body_html: str
    footnotes: List[FootnoteInfo] = field(default_factory=list)
    images: List[ImageInfo] = field(default_factory=list)
    has_equations: bool = False
    metadata: Dict[str, str] = field(default_factory=dict)

