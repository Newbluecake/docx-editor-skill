"""CLI entry point for docx tool.

Usage:
    docx <subcommand> <file> [options]

Each invocation: open file -> build registry -> execute -> save (if mutating) -> output Markdown
"""

import argparse
import sys
import logging

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Subcommand handlers
# ---------------------------------------------------------------------------
# Each handler receives (session, args) where session is the CLISession
# and args is the argparse Namespace. It returns the Markdown output string.
# Mutating commands are auto-saved after execution.


# -- Content ----------------------------------------------------------------

def _cmd_read(session, args):
    from docx_editor_skill.tools.content_tools import docx_read_content
    return docx_read_content(
        max_paragraphs=args.max_paragraphs,
        start_from=args.start_from,
        include_tables=args.include_tables,
        include_ids=True,
        start_element_id=args.start_element_id,
        max_tables=args.max_tables,
        table_mode=args.table_mode,
    )


def _cmd_find(session, args):
    from docx_editor_skill.tools.content_tools import docx_find_paragraphs
    return docx_find_paragraphs(
        query=args.query,
        max_results=args.max_results,
        return_context=args.context,
        case_sensitive=args.case_sensitive,
        context_span=args.context_span,
    )


def _cmd_structure(session, args):
    from docx_editor_skill.tools.content_tools import docx_extract_template_structure
    return docx_extract_template_structure(
        include_content=args.include_content,
        max_depth=args.max_depth,
        max_items_per_type=args.max_items,
    )


def _cmd_summary(session, args):
    from docx_editor_skill.tools.composite_tools import docx_get_structure_summary
    return docx_get_structure_summary(
        max_headings=args.max_headings,
        max_tables=args.max_tables,
        max_paragraphs=args.max_paragraphs,
        include_content=args.include_content,
    )


# -- Paragraph --------------------------------------------------------------

def _cmd_insert_paragraph(session, args):
    from docx_editor_skill.tools.paragraph_tools import docx_insert_paragraph
    return docx_insert_paragraph(
        text=args.text,
        position=args.position,
        style=args.style,
    )


def _cmd_insert_heading(session, args):
    from docx_editor_skill.tools.paragraph_tools import docx_insert_heading
    return docx_insert_heading(
        text=args.text,
        position=args.position,
        level=args.level,
    )


def _cmd_update_paragraph(session, args):
    from docx_editor_skill.tools.paragraph_tools import docx_update_paragraph_text
    return docx_update_paragraph_text(
        paragraph_id=args.paragraph_id,
        new_text=args.new_text,
    )


def _cmd_copy_paragraph(session, args):
    from docx_editor_skill.tools.paragraph_tools import docx_copy_paragraph
    return docx_copy_paragraph(
        paragraph_id=args.paragraph_id,
        position=args.position,
    )


def _cmd_delete(session, args):
    from docx_editor_skill.tools.paragraph_tools import docx_delete
    return docx_delete(element_id=args.element_id)


def _cmd_insert_page_break(session, args):
    from docx_editor_skill.tools.paragraph_tools import docx_insert_page_break
    return docx_insert_page_break(position=args.position)


# -- Run --------------------------------------------------------------------

def _cmd_insert_run(session, args):
    from docx_editor_skill.tools.run_tools import docx_insert_run
    return docx_insert_run(text=args.text, position=args.position)


def _cmd_update_run(session, args):
    from docx_editor_skill.tools.run_tools import docx_update_run_text
    return docx_update_run_text(run_id=args.run_id, new_text=args.new_text)


def _cmd_set_font(session, args):
    from docx_editor_skill.tools.run_tools import docx_set_font
    return docx_set_font(
        run_id=args.run_id,
        size=args.size,
        bold=args.bold,
        italic=args.italic,
        color_hex=args.color,
    )


# -- Table ------------------------------------------------------------------

def _cmd_insert_table(session, args):
    from docx_editor_skill.tools.table_tools import docx_insert_table
    return docx_insert_table(rows=args.rows, cols=args.cols, position=args.position)


def _cmd_get_table(session, args):
    from docx_editor_skill.tools.table_tools import docx_get_table
    return docx_get_table(index=args.index)


def _cmd_list_tables(session, args):
    from docx_editor_skill.tools.table_tools import docx_list_tables
    return docx_list_tables(
        max_results=args.max_results,
        start_element_id=args.start_element_id,
    )


def _cmd_find_table(session, args):
    from docx_editor_skill.tools.table_tools import docx_find_table
    return docx_find_table(
        text=args.text,
        max_results=args.max_results,
        start_element_id=args.start_element_id,
        return_structure=args.return_structure,
    )


def _cmd_get_cell(session, args):
    from docx_editor_skill.tools.table_tools import docx_get_cell
    return docx_get_cell(table_id=args.table_id, row=args.row, col=args.col)


def _cmd_insert_cell_text(session, args):
    from docx_editor_skill.tools.table_tools import docx_insert_paragraph_to_cell
    return docx_insert_paragraph_to_cell(text=args.text, position=args.position)


def _cmd_insert_table_row(session, args):
    from docx_editor_skill.tools.table_tools import docx_insert_table_row
    return docx_insert_table_row(position=args.position)


def _cmd_insert_table_col(session, args):
    from docx_editor_skill.tools.table_tools import docx_insert_table_col
    return docx_insert_table_col(position=args.position)


def _cmd_insert_row_at(session, args):
    from docx_editor_skill.tools.table_rowcol_tools import docx_insert_row_at
    return docx_insert_row_at(
        table_id=args.table_id,
        position=args.position,
        row_index=args.row_index,
        copy_format=args.copy_format,
    )


def _cmd_insert_col_at(session, args):
    from docx_editor_skill.tools.table_rowcol_tools import docx_insert_col_at
    return docx_insert_col_at(
        table_id=args.table_id,
        position=args.position,
        col_index=args.col_index,
        copy_format=args.copy_format,
    )


def _cmd_delete_row(session, args):
    from docx_editor_skill.tools.table_rowcol_tools import docx_delete_row
    return docx_delete_row(
        table_id=args.table_id,
        row_index=args.row_index,
        row_id=args.row_id,
    )


def _cmd_delete_col(session, args):
    from docx_editor_skill.tools.table_rowcol_tools import docx_delete_col
    return docx_delete_col(
        table_id=args.table_id,
        col_index=args.col_index,
        col_id=args.col_id,
    )


def _cmd_fill_table(session, args):
    from docx_editor_skill.tools.table_tools import docx_fill_table
    return docx_fill_table(
        data=args.data,
        table_id=args.table_id,
        start_row=args.start_row,
        preserve_formatting=args.preserve_formatting,
    )


def _cmd_copy_table(session, args):
    from docx_editor_skill.tools.table_tools import docx_copy_table
    return docx_copy_table(table_id=args.table_id, position=args.position)


def _cmd_table_structure(session, args):
    from docx_editor_skill.tools.table_tools import docx_get_table_structure
    return docx_get_table_structure(table_id=args.table_id)


# -- Format -----------------------------------------------------------------

def _cmd_set_alignment(session, args):
    from docx_editor_skill.tools.format_tools import docx_set_alignment
    return docx_set_alignment(paragraph_id=args.paragraph_id, alignment=args.alignment)


def _cmd_set_properties(session, args):
    from docx_editor_skill.tools.format_tools import docx_set_properties
    return docx_set_properties(properties=args.properties, element_id=args.element_id)


def _cmd_format_copy(session, args):
    from docx_editor_skill.tools.format_tools import docx_format_copy
    return docx_format_copy(source_id=args.source_id, target_id=args.target_id)


def _cmd_set_margins(session, args):
    from docx_editor_skill.tools.format_tools import docx_set_margins
    return docx_set_margins(
        top=args.top, bottom=args.bottom, left=args.left, right=args.right,
    )


def _cmd_extract_format(session, args):
    from docx_editor_skill.tools.format_tools import docx_extract_format_template
    return docx_extract_format_template(element_id=args.element_id)


def _cmd_apply_format(session, args):
    from docx_editor_skill.tools.format_tools import docx_apply_format_template
    return docx_apply_format_template(element_id=args.element_id, template_json=args.template_json)


# -- Advanced ---------------------------------------------------------------

def _cmd_replace_text(session, args):
    from docx_editor_skill.tools.advanced_tools import docx_replace_text
    return docx_replace_text(
        old_text=args.old_text, new_text=args.new_text, scope_id=args.scope_id,
    )


def _cmd_batch_replace(session, args):
    from docx_editor_skill.tools.advanced_tools import docx_batch_replace_text
    return docx_batch_replace_text(
        replacements_json=args.replacements_json, scope_id=args.scope_id,
    )


def _cmd_insert_image(session, args):
    from docx_editor_skill.tools.advanced_tools import docx_insert_image
    return docx_insert_image(
        image_path=args.image_path,
        position=args.position,
        width=args.width,
        height=args.height,
    )


# -- Composite --------------------------------------------------------------

def _cmd_insert_formatted(session, args):
    from docx_editor_skill.tools.composite_tools import docx_insert_formatted_paragraph
    return docx_insert_formatted_paragraph(
        text=args.text,
        position=args.position,
        bold=args.bold,
        italic=args.italic,
        size=args.size,
        color_hex=args.color,
        alignment=args.alignment,
        style=args.style,
    )


def _cmd_quick_edit(session, args):
    from docx_editor_skill.tools.composite_tools import docx_quick_edit
    return docx_quick_edit(
        search_text=args.search_text,
        new_text=args.new_text,
        bold=args.bold,
        italic=args.italic,
        size=args.size,
        color_hex=args.color,
    )


def _cmd_smart_fill(session, args):
    from docx_editor_skill.tools.composite_tools import docx_smart_fill_table
    return docx_smart_fill_table(
        table_identifier=args.table_identifier,
        data=args.data,
        has_header=args.has_header,
        auto_resize=args.auto_resize,
        preserve_formatting=args.preserve_formatting,
    )


def _cmd_format_range(session, args):
    from docx_editor_skill.tools.composite_tools import docx_format_range
    return docx_format_range(
        start_text=args.start_text,
        end_text=args.end_text,
        bold=args.bold,
        italic=args.italic,
        size=args.size,
        color_hex=args.color,
    )


# -- Copy -------------------------------------------------------------------

def _cmd_copy_range(session, args):
    from docx_editor_skill.tools.copy_tools import docx_copy_elements_range
    return docx_copy_elements_range(
        start_id=args.start_id, end_id=args.end_id, position=args.position,
    )


def _cmd_element_source(session, args):
    from docx_editor_skill.tools.copy_tools import docx_get_element_source
    return docx_get_element_source(element_id=args.element_id)


# -- File -------------------------------------------------------------------

def _cmd_save(session, args):
    from docx_editor_skill.tools.session_tools import docx_save
    return docx_save(file_path=args.output or session.file_path)


def _cmd_context(session, args):
    from docx_editor_skill.tools.session_tools import docx_get_context
    return docx_get_context()


# -- Preview ----------------------------------------------------------------

def _cmd_preview(session, args):
    from docx_editor_skill.tools.preview_tools import preview
    image_paths = preview(session.file_path, pages=args.pages, dpi=args.dpi)
    if not image_paths:
        return "No pages rendered."
    return "\n".join(image_paths)


def _cmd_preview_cleanup(session, args):
    from docx_editor_skill.tools.preview_tools import cleanup
    cleanup()
    return "Preview files cleaned up."


# ---------------------------------------------------------------------------
# Commands which mutate the document (auto-save after execution)
# ---------------------------------------------------------------------------
MUTATING_COMMANDS = {
    "insert-paragraph", "insert-heading", "update-paragraph", "copy-paragraph",
    "delete", "insert-page-break",
    "insert-run", "update-run", "set-font",
    "insert-table", "insert-cell-text", "insert-table-row", "insert-table-col",
    "insert-row-at", "insert-col-at", "delete-row", "delete-col",
    "fill-table", "copy-table",
    "set-alignment", "set-properties", "format-copy", "set-margins", "apply-format",
    "replace-text", "batch-replace", "insert-image",
    "insert-formatted", "quick-edit", "smart-fill", "format-range",
    "copy-range",
}

# Commands that don't need a file
NO_FILE_COMMANDS = {"preview-cleanup"}


# ---------------------------------------------------------------------------
# Argparse builder
# ---------------------------------------------------------------------------

def _add_bool_arg(parser, name, help_text):
    """Add a --flag / --no-flag boolean argument."""
    parser.add_argument(f"--{name}", action="store_true", default=None, help=help_text)
    parser.add_argument(f"--no-{name}", action="store_false", dest=name)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="docx",
        description="Word document manipulation CLI",
    )
    parser.add_argument("--no-save", action="store_true",
                        help="Skip auto-save for mutating commands")
    sub = parser.add_subparsers(dest="command", required=True)

    # --- Content ---
    p = sub.add_parser("read", help="Read document content")
    p.add_argument("file")
    p.add_argument("--max-paragraphs", type=int, default=None)
    p.add_argument("--start-from", type=int, default=0)
    p.add_argument("--include-tables", action="store_true")
    p.add_argument("--start-element-id", default=None)
    p.add_argument("--max-tables", type=int, default=None)
    p.add_argument("--table-mode", default="text")

    p = sub.add_parser("find", help="Find paragraphs containing text")
    p.add_argument("file")
    p.add_argument("query")
    p.add_argument("--max-results", type=int, default=10)
    p.add_argument("--context", action="store_true")
    p.add_argument("--case-sensitive", action="store_true")
    p.add_argument("--context-span", type=int, default=0)

    p = sub.add_parser("structure", help="Extract full template structure")
    p.add_argument("file")
    p.add_argument("--include-content", action="store_true", default=True)
    p.add_argument("--no-content", action="store_false", dest="include_content")
    p.add_argument("--max-depth", type=int, default=None)
    p.add_argument("--max-items", default=None, help="JSON dict limiting items per type")

    p = sub.add_parser("summary", help="Lightweight structure summary")
    p.add_argument("file")
    p.add_argument("--max-headings", type=int, default=10)
    p.add_argument("--max-tables", type=int, default=5)
    p.add_argument("--max-paragraphs", type=int, default=0)
    p.add_argument("--include-content", action="store_true")

    # --- Paragraph ---
    p = sub.add_parser("insert-paragraph", help="Insert a paragraph")
    p.add_argument("file")
    p.add_argument("text")
    p.add_argument("--position", required=True)
    p.add_argument("--style", default=None)

    p = sub.add_parser("insert-heading", help="Insert a heading")
    p.add_argument("file")
    p.add_argument("text")
    p.add_argument("--position", required=True)
    p.add_argument("--level", type=int, default=1)

    p = sub.add_parser("update-paragraph", help="Update paragraph text")
    p.add_argument("file")
    p.add_argument("paragraph_id")
    p.add_argument("new_text")

    p = sub.add_parser("copy-paragraph", help="Copy a paragraph")
    p.add_argument("file")
    p.add_argument("paragraph_id")
    p.add_argument("--position", required=True)

    p = sub.add_parser("delete", help="Delete an element")
    p.add_argument("file")
    p.add_argument("element_id")

    p = sub.add_parser("insert-page-break", help="Insert a page break")
    p.add_argument("file")
    p.add_argument("--position", required=True)

    # --- Run ---
    p = sub.add_parser("insert-run", help="Insert a text run")
    p.add_argument("file")
    p.add_argument("text")
    p.add_argument("--position", required=True)

    p = sub.add_parser("update-run", help="Update run text")
    p.add_argument("file")
    p.add_argument("run_id")
    p.add_argument("new_text")

    p = sub.add_parser("set-font", help="Set font properties on a run")
    p.add_argument("file")
    p.add_argument("run_id")
    p.add_argument("--size", type=float, default=None)
    p.add_argument("--bold", action="store_true", default=None)
    p.add_argument("--no-bold", action="store_false", dest="bold")
    p.add_argument("--italic", action="store_true", default=None)
    p.add_argument("--no-italic", action="store_false", dest="italic")
    p.add_argument("--color", default=None, help="Hex color e.g. FF0000")

    # --- Table ---
    p = sub.add_parser("insert-table", help="Create a table")
    p.add_argument("file")
    p.add_argument("--rows", type=int, required=True)
    p.add_argument("--cols", type=int, required=True)
    p.add_argument("--position", required=True)

    p = sub.add_parser("get-table", help="Get table by index")
    p.add_argument("file")
    p.add_argument("index", type=int)

    p = sub.add_parser("list-tables", help="List tables")
    p.add_argument("file")
    p.add_argument("--max-results", type=int, default=50)
    p.add_argument("--start-element-id", default=None)

    p = sub.add_parser("find-table", help="Find table by text")
    p.add_argument("file")
    p.add_argument("text")
    p.add_argument("--max-results", type=int, default=1)
    p.add_argument("--start-element-id", default=None)
    p.add_argument("--return-structure", action="store_true")

    p = sub.add_parser("get-cell", help="Get a table cell")
    p.add_argument("file")
    p.add_argument("table_id")
    p.add_argument("--row", type=int, required=True)
    p.add_argument("--col", type=int, required=True)

    p = sub.add_parser("insert-cell-text", help="Insert text in table cell")
    p.add_argument("file")
    p.add_argument("text")
    p.add_argument("--position", required=True)

    p = sub.add_parser("insert-table-row", help="Add row to table")
    p.add_argument("file")
    p.add_argument("--position", required=True)

    p = sub.add_parser("insert-table-col", help="Add column to table")
    p.add_argument("file")
    p.add_argument("--position", required=True)

    p = sub.add_parser("insert-row-at", help="Insert row at position")
    p.add_argument("file")
    p.add_argument("table_id")
    p.add_argument("--position", required=True)
    p.add_argument("--row-index", type=int, default=None)
    p.add_argument("--copy-format", action="store_true")

    p = sub.add_parser("insert-col-at", help="Insert column at position")
    p.add_argument("file")
    p.add_argument("table_id")
    p.add_argument("--position", required=True)
    p.add_argument("--col-index", type=int, default=None)
    p.add_argument("--copy-format", action="store_true")

    p = sub.add_parser("delete-row", help="Delete a table row")
    p.add_argument("file")
    p.add_argument("table_id")
    p.add_argument("--row-index", type=int, default=None)
    p.add_argument("--row-id", default=None)

    p = sub.add_parser("delete-col", help="Delete a table column")
    p.add_argument("file")
    p.add_argument("table_id")
    p.add_argument("--col-index", type=int, default=None)
    p.add_argument("--col-id", default=None)

    p = sub.add_parser("fill-table", help="Batch fill table data")
    p.add_argument("file")
    p.add_argument("data", help="JSON 2D array string")
    p.add_argument("--table-id", default=None)
    p.add_argument("--start-row", type=int, default=0)
    p.add_argument("--preserve-formatting", action="store_true", default=True)
    p.add_argument("--no-preserve-formatting", action="store_false", dest="preserve_formatting")

    p = sub.add_parser("copy-table", help="Copy a table")
    p.add_argument("file")
    p.add_argument("table_id")
    p.add_argument("--position", required=True)

    p = sub.add_parser("table-structure", help="Show table structure")
    p.add_argument("file")
    p.add_argument("table_id")

    # --- Format ---
    p = sub.add_parser("set-alignment", help="Set paragraph alignment")
    p.add_argument("file")
    p.add_argument("paragraph_id")
    p.add_argument("alignment", choices=["left", "center", "right", "justify"])

    p = sub.add_parser("set-properties", help="Set element properties (JSON)")
    p.add_argument("file")
    p.add_argument("properties", help="JSON properties string")
    p.add_argument("--element-id", default=None)

    p = sub.add_parser("format-copy", help="Copy format from source to target")
    p.add_argument("file")
    p.add_argument("source_id")
    p.add_argument("target_id")

    p = sub.add_parser("set-margins", help="Set page margins (inches)")
    p.add_argument("file")
    p.add_argument("--top", type=float, default=None)
    p.add_argument("--bottom", type=float, default=None)
    p.add_argument("--left", type=float, default=None)
    p.add_argument("--right", type=float, default=None)

    p = sub.add_parser("extract-format", help="Extract format template from element")
    p.add_argument("file")
    p.add_argument("element_id")

    p = sub.add_parser("apply-format", help="Apply format template to element")
    p.add_argument("file")
    p.add_argument("element_id")
    p.add_argument("template_json", help="JSON format template string")

    # --- Advanced ---
    p = sub.add_parser("replace-text", help="Replace text in document")
    p.add_argument("file")
    p.add_argument("old_text")
    p.add_argument("new_text")
    p.add_argument("--scope-id", default=None)

    p = sub.add_parser("batch-replace", help="Batch replace text")
    p.add_argument("file")
    p.add_argument("replacements_json", help='JSON: {"old": "new", ...}')
    p.add_argument("--scope-id", default=None)

    p = sub.add_parser("insert-image", help="Insert an image")
    p.add_argument("file")
    p.add_argument("image_path")
    p.add_argument("--position", required=True)
    p.add_argument("--width", type=float, default=None)
    p.add_argument("--height", type=float, default=None)

    # --- Composite ---
    p = sub.add_parser("insert-formatted", help="Insert formatted paragraph")
    p.add_argument("file")
    p.add_argument("text")
    p.add_argument("--position", required=True)
    p.add_argument("--bold", action="store_true", default=False)
    p.add_argument("--italic", action="store_true", default=False)
    p.add_argument("--size", type=float, default=None)
    p.add_argument("--color", default=None)
    p.add_argument("--alignment", default=None)
    p.add_argument("--style", default=None)

    p = sub.add_parser("quick-edit", help="Find and edit paragraphs")
    p.add_argument("file")
    p.add_argument("search_text")
    p.add_argument("--new-text", default=None)
    p.add_argument("--bold", action="store_true", default=None)
    p.add_argument("--no-bold", action="store_false", dest="bold")
    p.add_argument("--italic", action="store_true", default=None)
    p.add_argument("--no-italic", action="store_false", dest="italic")
    p.add_argument("--size", type=float, default=None)
    p.add_argument("--color", default=None)

    p = sub.add_parser("smart-fill", help="Smart fill a table")
    p.add_argument("file")
    p.add_argument("table_identifier")
    p.add_argument("data", help="JSON 2D array")
    p.add_argument("--has-header", action="store_true", default=True)
    p.add_argument("--no-header", action="store_false", dest="has_header")
    p.add_argument("--auto-resize", action="store_true", default=True)
    p.add_argument("--no-auto-resize", action="store_false", dest="auto_resize")
    p.add_argument("--preserve-formatting", action="store_true", default=False)

    p = sub.add_parser("format-range", help="Format a range of paragraphs")
    p.add_argument("file")
    p.add_argument("start_text")
    p.add_argument("end_text")
    p.add_argument("--bold", action="store_true", default=None)
    p.add_argument("--no-bold", action="store_false", dest="bold")
    p.add_argument("--italic", action="store_true", default=None)
    p.add_argument("--no-italic", action="store_false", dest="italic")
    p.add_argument("--size", type=float, default=None)
    p.add_argument("--color", default=None)

    # --- Copy ---
    p = sub.add_parser("copy-range", help="Copy elements range")
    p.add_argument("file")
    p.add_argument("start_id")
    p.add_argument("end_id")
    p.add_argument("--position", required=True)

    p = sub.add_parser("element-source", help="Get element source metadata")
    p.add_argument("file")
    p.add_argument("element_id")

    # --- File ---
    p = sub.add_parser("save", help="Save document (explicit save-as)")
    p.add_argument("file")
    p.add_argument("--output", "-o", default=None, help="Output path (default: overwrite input)")

    p = sub.add_parser("context", help="Show session context")
    p.add_argument("file")

    # --- Preview ---
    p = sub.add_parser("preview", help="Preview document as images")
    p.add_argument("file")
    p.add_argument("--pages", "-p", default="1")
    p.add_argument("--dpi", "-d", type=int, default=200)

    p = sub.add_parser("preview-cleanup", help="Remove preview temp files")

    return parser


# ---------------------------------------------------------------------------
# Command dispatch table
# ---------------------------------------------------------------------------
DISPATCH = {
    # Content
    "read": _cmd_read,
    "find": _cmd_find,
    "structure": _cmd_structure,
    "summary": _cmd_summary,
    # Paragraph
    "insert-paragraph": _cmd_insert_paragraph,
    "insert-heading": _cmd_insert_heading,
    "update-paragraph": _cmd_update_paragraph,
    "copy-paragraph": _cmd_copy_paragraph,
    "delete": _cmd_delete,
    "insert-page-break": _cmd_insert_page_break,
    # Run
    "insert-run": _cmd_insert_run,
    "update-run": _cmd_update_run,
    "set-font": _cmd_set_font,
    # Table
    "insert-table": _cmd_insert_table,
    "get-table": _cmd_get_table,
    "list-tables": _cmd_list_tables,
    "find-table": _cmd_find_table,
    "get-cell": _cmd_get_cell,
    "insert-cell-text": _cmd_insert_cell_text,
    "insert-table-row": _cmd_insert_table_row,
    "insert-table-col": _cmd_insert_table_col,
    "insert-row-at": _cmd_insert_row_at,
    "insert-col-at": _cmd_insert_col_at,
    "delete-row": _cmd_delete_row,
    "delete-col": _cmd_delete_col,
    "fill-table": _cmd_fill_table,
    "copy-table": _cmd_copy_table,
    "table-structure": _cmd_table_structure,
    # Format
    "set-alignment": _cmd_set_alignment,
    "set-properties": _cmd_set_properties,
    "format-copy": _cmd_format_copy,
    "set-margins": _cmd_set_margins,
    "extract-format": _cmd_extract_format,
    "apply-format": _cmd_apply_format,
    # Advanced
    "replace-text": _cmd_replace_text,
    "batch-replace": _cmd_batch_replace,
    "insert-image": _cmd_insert_image,
    # Composite
    "insert-formatted": _cmd_insert_formatted,
    "quick-edit": _cmd_quick_edit,
    "smart-fill": _cmd_smart_fill,
    "format-range": _cmd_format_range,
    # Copy
    "copy-range": _cmd_copy_range,
    "element-source": _cmd_element_source,
    # File
    "save": _cmd_save,
    "context": _cmd_context,
    # Preview
    "preview": _cmd_preview,
    "preview-cleanup": _cmd_preview_cleanup,
}


# ---------------------------------------------------------------------------
# Main entry
# ---------------------------------------------------------------------------

def cli_main():
    parser = build_parser()
    args = parser.parse_args()

    command = args.command
    handler = DISPATCH.get(command)
    if not handler:
        print(f"Unknown command: {command}", file=sys.stderr)
        sys.exit(1)

    # Commands that don't need a file
    if command in NO_FILE_COMMANDS:
        result = handler(None, args)
        print(result)
        return

    # All other commands need a file
    file_path = args.file

    from docx_editor_skill.core.cli_session import open_cli_session

    try:
        session = open_cli_session(file_path)
    except Exception as e:
        print(f"Error opening {file_path}: {e}", file=sys.stderr)
        sys.exit(1)

    # Execute command
    try:
        result = handler(session, args)
    except Exception as e:
        logger.exception(f"Command {command} failed")
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    # Auto-save for mutating commands
    if command in MUTATING_COMMANDS and not args.no_save:
        try:
            session.document.save(session.file_path)
        except Exception as e:
            print(f"Warning: auto-save failed: {e}", file=sys.stderr)

    # Output
    print(result)


if __name__ == "__main__":
    cli_main()
