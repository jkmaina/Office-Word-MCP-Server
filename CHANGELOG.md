# Changelog

All notable changes to this project will be documented in this file.

## [1.8.0] - 2025-05-29
### Added
- Export tools (Phase 6):
  - `to_epub(filename, output_filename, metadata, toc)` to convert .docx to EPUB via Pandoc.
  - `to_pdf(filename, output_filename, pdf_engine)` to convert .docx to PDF via Pandoc.
  - `set_core_properties(filename, title, author, subject, keywords)` to inject document metadata.

## [1.7.0] - 2025-05-29
### Added
Cross-reference tools (Phase 5):
  - `bookmark(filename, bookmark_name)` to mark a location for hyperlinking.
  - `insert_hyperlink(filename, bookmark_name, display_text)` to add an internal link to a bookmark.
Caption & list-generation tools (Phase 5):
  - `insert_caption(filename, object_type, text, style)` to insert numbered captions for figures/tables.
  - `generate_list_of_figures(filename, heading_style)` to insert a List of Figures field.
  - `generate_list_of_tables(filename, heading_style)` to insert a List of Tables field.

## [1.6.0] - 2025-05-29
### Added
Book & front-matter tools (Phase 4):
  - `add_title_page(filename, title, subtitle, author, date)` to create a standalone title page doc.
  - `add_copyright_page(filename, text)` to append a copyright/legal page.
  - `add_front_matter(filename, section, text)` to insert preface/acknowledgment/dedication sections.
Navigation tools (Phase 4):
  - `insert_toc_placeholder(filename, heading_style)` to inject a live TOC field placeholder.

## [1.5.0] - 2025-05-29
### Added
Code block insertion tool (Phase 3):
  - `add_code_block(filename, code_text, language, style)` inserts a monospaced, shaded code block as a one-cell table.

## [1.4.0] - 2025-05-29
### Added
Header & footer tools (Phase 2):
  - `insert_header(filename, text, alignment, fields)` to inject headers with placeholders (e.g. `{pagenum}`).
  - `insert_footer(filename, text, alignment, fields)` to inject footers with placeholders.
Chapter tools (Phase 2):
  - `new_chapter(filename, title, style)` auto-numbers chapters by counting existing headings and inserts a section break.

## [1.3.0] - 2025-05-29
### Added
Template & layout tools:
  - `apply_template(template_path, destination_path)` to bootstrap a new document from a .docx/.dotx template.
  - `set_page_size(width, height, margins)` to set page dimensions (in inches) and optional margins on all sections.
  - `add_section_break(break_type)` to insert a section break (`nextPage`, `evenPage`, or `oddPage`).

## [1.2.0] - 2025-05-29
### Added
- Initial test suite for core tool modules (smoke-imports and interface checks).
- Continuous Integration workflow (GitHub Actions) to run pytest and build distributions via Hatch.
- Semantic Versioning policy: minor versions add features, patch versions for bugfixes.