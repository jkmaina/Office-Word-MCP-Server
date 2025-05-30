"""
Main entry point for the Word Document MCP Server.
Acts as the central controller for the MCP server that handles Word document operations.
"""

import os
import sys
from mcp.server.fastmcp import FastMCP
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
)



# Initialize FastMCP server
mcp = FastMCP("word-document-server")

def register_tools():
    """Register all tools with the MCP server."""
    # Document tools (create, copy, info, etc.)
    mcp.tool()(document_tools.create_document)
    mcp.tool()(document_tools.copy_document)
    mcp.tool()(document_tools.get_document_info)
    mcp.tool()(document_tools.get_document_text)
    mcp.tool()(document_tools.get_document_outline)
    mcp.tool()(document_tools.list_available_documents)
    
    # Content tools (paragraphs, headings, tables, etc.)
    mcp.tool()(content_tools.add_paragraph)
    mcp.tool()(content_tools.add_heading)
    mcp.tool()(content_tools.add_picture)
    mcp.tool()(content_tools.add_table)
    mcp.tool()(content_tools.add_page_break)
    mcp.tool()(content_tools.delete_paragraph)
    mcp.tool()(content_tools.search_and_replace)
    
    # Format tools (styling, text formatting, etc.)
    mcp.tool()(format_tools.create_custom_style)
    mcp.tool()(format_tools.format_text)
    mcp.tool()(format_tools.format_table)
    
    # Protection tools
    mcp.tool()(protection_tools.protect_document)
    mcp.tool()(protection_tools.unprotect_document)
    
    # Footnote tools
    mcp.tool()(footnote_tools.add_footnote_to_document)
    mcp.tool()(footnote_tools.add_endnote_to_document)
    mcp.tool()(footnote_tools.convert_footnotes_to_endnotes_in_document)
    mcp.tool()(footnote_tools.customize_footnote_style)
    
    # Extended document tools
    mcp.tool()(extended_document_tools.get_paragraph_text_from_document)
    mcp.tool()(extended_document_tools.find_text_in_document)
    mcp.tool()(extended_document_tools.convert_to_pdf)
    
    # Template & layout tools (Phase 1)
    mcp.tool()(template_tools.apply_template)
    mcp.tool()(template_tools.set_page_size)
    mcp.tool()(template_tools.add_section_break)
    # Header & footer tools (Phase 2)
    mcp.tool()(header_tools.insert_header)
    mcp.tool()(header_tools.insert_footer)
    # Chapter tools (Phase 2)
    mcp.tool()(chapter_tools.new_chapter)
    # Code block tools (Phase 3)
    mcp.tool()(code_tools.add_code_block)
    # Front/Back matter tools (Phase 4)
    mcp.tool()(book_tools.add_title_page)
    mcp.tool()(book_tools.add_copyright_page)
    mcp.tool()(book_tools.add_front_matter)
    # Navigation tools (Phase 4)
    mcp.tool()(navigation_tools.insert_toc_placeholder)
    # Cross-reference tools (Phase 5)
    mcp.tool()(navigation_tools.bookmark)
    mcp.tool()(navigation_tools.insert_hyperlink)
    # Caption and list tools (Phase 5)
    mcp.tool()(caption_tools.insert_caption)
    mcp.tool()(caption_tools.generate_list_of_figures)
    mcp.tool()(caption_tools.generate_list_of_tables)


def run_server():
    """Run the Word Document MCP Server."""
    # Register all tools
    register_tools()
    
    # Run the server
    mcp.run(transport='stdio')
    return mcp

if __name__ == "__main__":
    run_server()
