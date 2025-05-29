import pytest

from word_document_server.tools import (
    document_tools,
    content_tools,
    format_tools,
    protection_tools,
    footnote_tools,
    extended_document_tools,
)

def test_document_tools_has_core_functions():
    assert hasattr(document_tools, 'create_document')
    assert hasattr(document_tools, 'copy_document')
    assert hasattr(document_tools, 'get_document_info')

def test_content_tools_has_core_functions():
    assert hasattr(content_tools, 'add_paragraph')
    assert hasattr(content_tools, 'add_heading')
    assert hasattr(content_tools, 'add_table')

def test_format_tools_has_core_functions():
    assert hasattr(format_tools, 'format_text')
    assert hasattr(format_tools, 'create_custom_style')

def test_protection_tools_has_core_functions():
    assert hasattr(protection_tools, 'protect_document')
    assert hasattr(protection_tools, 'unprotect_document')

def test_footnote_tools_has_core_functions():
    assert hasattr(footnote_tools, 'add_footnote_to_document')
    assert hasattr(footnote_tools, 'convert_footnotes_to_endnotes_in_document')

def test_extended_document_tools_has_core_functions():
    assert hasattr(extended_document_tools, 'convert_to_pdf')
    assert hasattr(extended_document_tools, 'find_text_in_document')