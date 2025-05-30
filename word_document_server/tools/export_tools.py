"""
Export tools for Word Document Server.

These tools convert .docx to EPUB or PDF via Pandoc and inject document metadata.
"""
import os
import shutil
import subprocess
from typing import Optional, List, Dict

from docx import Document

from word_document_server.utils.file_utils import ensure_docx_extension, check_file_writeable


async def to_epub(
    filename: str,
    output_filename: Optional[str] = None,
    metadata: Optional[Dict[str, str]] = None,
    toc: bool = True,
) -> str:
    """Convert a Word document to EPUB using Pandoc.

    Args:
        filename: Path to the source .docx file
        output_filename: Optional path for the output .epub file
        metadata: Optional dict of metadata fields (title, author, etc.)
        toc: Whether to include a table of contents
    """
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    # Determine output filename
    if not output_filename:
        base, _ = os.path.splitext(filename)
        output_filename = base + '.epub'
    # Check writeability
    writeable, err = check_file_writeable(output_filename)
    if not writeable:
        return f"Cannot create EPUB: {err}"
    # Check Pandoc availability
    if not shutil.which('pandoc'):
        return "Pandoc is not installed or not in PATH."
    # Build command
    cmd = ['pandoc', filename, '-o', output_filename]
    if toc:
        cmd.append('--toc')
    if metadata:
        for key, value in metadata.items():
            cmd.extend(['--metadata', f'{key}={value}'])
    try:
        subprocess.run(cmd, check=True, capture_output=True)
        return f"EPUB created: {output_filename}"
    except subprocess.CalledProcessError as e:
        return f"Failed to create EPUB: {e.stderr.decode().strip()}"
    except Exception as e:
        return f"Error during EPUB conversion: {str(e)}"


async def to_pdf(
    filename: str,
    output_filename: Optional[str] = None,
    pdf_engine: str = 'xelatex',
) -> str:
    """Convert a Word document to PDF using Pandoc and a PDF engine.

    Args:
        filename: Path to the source .docx file
        output_filename: Optional path for the output .pdf file
        pdf_engine: PDF engine for Pandoc (e.g. xelatex, pdflatex)
    """
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    # Determine output filename
    if not output_filename:
        base, _ = os.path.splitext(filename)
        output_filename = base + '.pdf'
    # Check writeability
    writeable, err = check_file_writeable(output_filename)
    if not writeable:
        return f"Cannot create PDF: {err}"
    # Check Pandoc availability
    if not shutil.which('pandoc'):
        return "Pandoc is not installed or not in PATH."
    # Build command
    cmd = ['pandoc', filename, '-o', output_filename, '--pdf-engine', pdf_engine]
    try:
        subprocess.run(cmd, check=True, capture_output=True)
        return f"PDF created: {output_filename}"
    except subprocess.CalledProcessError as e:
        return f"Failed to create PDF: {e.stderr.decode().strip()}"
    except Exception as e:
        return f"Error during PDF conversion: {str(e)}"


async def set_core_properties(
    filename: str,
    title: Optional[str] = None,
    author: Optional[str] = None,
    subject: Optional[str] = None,
    keywords: Optional[List[str]] = None,
) -> str:
    """Set core document properties (metadata) on a .docx file.

    Args:
        filename: Path to the .docx file
        title: Document title
        author: Document author
        subject: Document subject
        keywords: List of keyword strings
    """
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    writable, err = check_file_writeable(filename)
    if not writable:
        return f"Cannot modify document: {err}"
    try:
        doc = Document(filename)
        props = doc.core_properties
        if title is not None:
            props.title = title
        if author is not None:
            props.author = author
        if subject is not None:
            props.subject = subject
        if keywords is not None:
            props.keywords = ', '.join(keywords)
        doc.save(filename)
        return f"Metadata updated on {filename}"
    except Exception as e:
        return f"Failed to set core properties: {e}"