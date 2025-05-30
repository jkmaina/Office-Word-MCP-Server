"""
Navigation tools for Word Document Server.

These tools insert table of contents placeholders and manage cross-document navigation.
"""
import os

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.shared import qn

from word_document_server.utils.file_utils import ensure_docx_extension, check_file_writeable


async def insert_toc_placeholder(
    filename: str,
    heading_style: str = 'TOC Heading',
) -> str:
    """Insert a Table of Contents field placeholder.

    Args:
        filename: Path to an existing Word document
        heading_style: Style name to apply to the TOC heading
    """
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    writeable, err = check_file_writeable(filename)
    if not writeable:
        return f"Cannot modify document: {err}"
    try:
        doc = Document(filename)
        # Insert heading for TOC
        toc_heading = doc.add_paragraph('Table of Contents')
        try:
            toc_heading.style = heading_style
        except Exception:
            pass
        # Insert the TOC field
        p = doc.add_paragraph()
        fld = OxmlElement('w:fldSimple')
        # Classic TOC field code (levels 1-3, hyperlinks, hide page numbers)
        fld.set(qn('w:instr'), r'TOC \o "1-3" \h \z \u')
        p._p.append(fld)
        doc.save(filename)
        return f"TOC placeholder inserted into {filename}"
    except Exception as e:
        return f"Failed to insert TOC placeholder: {e}"