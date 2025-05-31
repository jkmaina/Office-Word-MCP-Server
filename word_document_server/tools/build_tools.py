"""
Build orchestrator for Word Document Server.

This tool drives an end-to-end book assembly pipeline based on
a manifest of steps (JSON format).
"""
import os
import json

from word_document_server.tools.template_tools import apply_template
from word_document_server.tools.header_tools import insert_header, insert_footer
from word_document_server.tools.chapter_tools import new_chapter
from word_document_server.tools.book_tools import add_title_page, add_front_matter, add_copyright_page
from word_document_server.tools.navigation_tools import insert_toc_placeholder, bookmark, insert_hyperlink
from word_document_server.tools.caption_tools import insert_caption, generate_list_of_figures, generate_list_of_tables
from word_document_server.tools.export_tools import to_epub, to_pdf, set_core_properties


async def build_book(
    manifest_path: str,
) -> str:
    """Run a build pipeline using a JSON manifest file.

    The manifest JSON should have the format:
    {
      "steps": [
        {"tool": "apply_template", "args": { ... }},
        ...
      ]
    }

    Returns a JSON report of each step's result.
    """
    if not os.path.exists(manifest_path):
        return f"Manifest file not found: {manifest_path}"
    try:
        with open(manifest_path, 'r') as f:
            manifest = json.load(f)
        steps = manifest.get('steps', [])
        report = []
        # Map tool names to callables
        tool_map = {
            'apply_template': apply_template,
            'insert_header': insert_header,
            'insert_footer': insert_footer,
            'new_chapter': new_chapter,
            'add_title_page': add_title_page,
            'add_front_matter': add_front_matter,
            'add_copyright_page': add_copyright_page,
            'insert_toc_placeholder': insert_toc_placeholder,
            'bookmark': bookmark,
            'insert_hyperlink': insert_hyperlink,
            'insert_caption': insert_caption,
            'generate_list_of_figures': generate_list_of_figures,
            'generate_list_of_tables': generate_list_of_tables,
            'to_epub': to_epub,
            'to_pdf': to_pdf,
            'set_core_properties': set_core_properties,
        }
        for step in steps:
            tool = step.get('tool')
            args = step.get('args', {})
            func = tool_map.get(tool)
            if not func:
                result = f"Unknown tool: {tool}"
            else:
                try:
                    # Await the async tool call
                    res = await func(**args)
                    result = res
                except Exception as e:
                    result = f"Error running {tool}: {e}"
            report.append({'tool': tool, 'result': result})
        return json.dumps(report, indent=2)
    except Exception as e:
        return f"Failed to build book: {e}"