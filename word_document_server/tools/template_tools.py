"""
Template and layout tools for Word Document Server.

These tools handle applying document templates, setting page dimensions/margins,
and inserting section breaks for print-ready layouts.
"""
import os
from typing import Optional, Dict

from docx import Document
from docx.shared import Inches
from docx.enum.section import WD_SECTION_START

from word_document_server.utils.file_utils import (
    ensure_docx_extension,
    check_file_writeable,
    create_document_copy,
)

async def apply_template(template_path: str, destination_path: Optional[str] = None) -> str:
    """Create a new document based on a .docx/.dotx template.

    Args:
        template_path: Path to the template file (.docx or .dotx)
        destination_path: Optional path for the new document (will ensure .docx extension)
    """
    # Validate template file
    ext = os.path.splitext(template_path)[1].lower()
    if ext not in ['.docx', '.dotx']:
        return f"Template must be .docx or .dotx: {template_path}"
    if not os.path.exists(template_path):
        return f"Template file not found: {template_path}"

    # Determine destination
    if destination_path:
        destination_path = ensure_docx_extension(destination_path)
    # Copy template to create new document
    success, message, new_path = create_document_copy(template_path, destination_path)
    if success:
        return message
    return f"Failed to apply template: {message}"

async def set_page_size(
    filename: str,
    width: float,
    height: float,
    margins: Optional[Dict[str, float]] = None,
) -> str:
    """Set page size (in inches) and optional margins for all sections.

    Args:
        filename: Path to the Word document to modify
        width: Page width in inches
        height: Page height in inches
        margins: Optional dict of margins in inches with keys 'top', 'bottom', 'left',
                 'right', 'header', 'footer', 'gutter'
    """
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    is_writeable, err = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {err}"
    try:
        doc = Document(filename)
        for section in doc.sections:
            section.page_width = Inches(width)
            section.page_height = Inches(height)
            if margins:
                if 'top' in margins:
                    section.top_margin = Inches(margins['top'])
                if 'bottom' in margins:
                    section.bottom_margin = Inches(margins['bottom'])
                if 'left' in margins:
                    section.left_margin = Inches(margins['left'])
                if 'right' in margins:
                    section.right_margin = Inches(margins['right'])
                if 'header' in margins:
                    section.header_distance = Inches(margins['header'])
                if 'footer' in margins:
                    section.footer_distance = Inches(margins['footer'])
                if 'gutter' in margins:
                    section.gutter = Inches(margins['gutter'])
        doc.save(filename)
        return f"Page size set to {width}x{height} inches in {filename}"
    except Exception as e:
        return f"Failed to set page size: {str(e)}"

async def add_section_break(
    filename: str,
    break_type: str = "nextPage",
) -> str:
    """Insert a section break of the specified type.

    Args:
        filename: Path to the Word document to modify
        break_type: Type of section break: 'nextPage', 'evenPage', or 'oddPage'
    """
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    is_writeable, err = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {err}"
    types_map = {
        'nextPage': WD_SECTION_START.NEXT_PAGE,
        'evenPage': WD_SECTION_START.EVEN_PAGE,
        'oddPage': WD_SECTION_START.ODD_PAGE,
    }
    if break_type not in types_map:
        return f"Invalid break_type: {break_type}. Must be one of {list(types_map.keys())}"
    try:
        doc = Document(filename)
        # Attempt to insert a real section break
        try:
            doc.add_section(types_map[break_type])
        except Exception:
            # Fallback to page break if sections unsupported
            doc.add_page_break()
        doc.save(filename)
        return f"Section break ({break_type}) added to {filename}"
    except Exception as e:
        return f"Failed to add section break: {str(e)}"