import pytest

from word_document_server.tools import (
    document_tools,
    content_tools,
    format_tools,
    protection_tools,
    footnote_tools,
    extended_document_tools,
    template_tools,
    header_tools,
    chapter_tools,
    code_tools,
    book_tools,
    navigation_tools,
    caption_tools,
    export_tools,
    qa_tools,
    build_tools,
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
    
def test_template_tools_has_core_functions():
    assert hasattr(template_tools, 'apply_template')
    assert hasattr(template_tools, 'set_page_size')
    assert hasattr(template_tools, 'add_section_break')
    
def test_header_tools_has_core_functions():
    assert hasattr(header_tools, 'insert_header')
    assert hasattr(header_tools, 'insert_footer')

def test_chapter_tools_has_core_functions():
    assert hasattr(chapter_tools, 'new_chapter')
    
def test_code_tools_has_core_functions():
    assert hasattr(code_tools, 'add_code_block')
    
def test_book_tools_has_core_functions():
    assert hasattr(book_tools, 'add_title_page')
    assert hasattr(book_tools, 'add_copyright_page')
    assert hasattr(book_tools, 'add_front_matter')

def test_navigation_tools_has_core_functions():
    assert hasattr(navigation_tools, 'insert_toc_placeholder')
    assert hasattr(navigation_tools, 'bookmark')
    assert hasattr(navigation_tools, 'insert_hyperlink')

def test_caption_tools_has_core_functions():
    assert hasattr(caption_tools, 'insert_caption')
    assert hasattr(caption_tools, 'generate_list_of_figures')
    assert hasattr(caption_tools, 'generate_list_of_tables')
    
def test_export_tools_has_core_functions():
    assert hasattr(export_tools, 'to_epub')
    assert hasattr(export_tools, 'to_pdf')
    assert hasattr(export_tools, 'set_core_properties')
    
def test_qa_tools_has_core_functions():
    assert hasattr(qa_tools, 'extract_text')
    assert hasattr(qa_tools, 'check_sentence_length')
    assert hasattr(qa_tools, 'check_passive_voice')

def test_build_tools_has_core_functions():
    assert hasattr(build_tools, 'build_book')