# ============= setup_pandoc.py =============
# ============= setup_pandoc.py =============
import pypandoc
from pathlib import Path

# Download to your current project
project_dir = Path(__file__).parent  # Same folder as this script
pandoc_folder = project_dir / "pandoc_bin"

print(f"Downloading pandoc to: {pandoc_folder}")
pypandoc.download_pandoc(
    targetfolder=str(pandoc_folder),
    version='3.1.9'
)
print("Done! Pandoc is in:", pandoc_folder)
exit(0)
"""Setup script to download pandoc"""
import pypandoc

# Download pandoc executable
print("Downloading pandoc...")
pypandoc.download_pandoc(
    targetfolder="C:/pandoc/",  # Or any folder you prefer
    version='3.1.9'  # Latest stable version
)
print("Pandoc downloaded successfully!")

# Add to PATH or use pypandoc with specific path
import os
os.environ['PATH'] = r'C:\pandoc;' + os.environ['PATH']