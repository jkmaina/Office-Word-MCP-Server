# Changelog

All notable changes to this project will be documented in this file.

## [1.4.0] - 2025-05-29
### Added
- Header & footer tools (Phase 2):
  - `insert_header(filename, text, alignment, fields)` to inject headers with placeholders (e.g. `{pagenum}`).
  - `insert_footer(filename, text, alignment, fields)` to inject footers with placeholders.
- Chapter tools (Phase 2):
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