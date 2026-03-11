"""Tool function registry — pure imports, no MCP dependency."""

from .session_tools import docx_save, docx_get_context
from .content_tools import docx_read_content, docx_find_paragraphs, docx_extract_template_structure
from .paragraph_tools import (
    docx_insert_paragraph, docx_insert_heading, docx_update_paragraph_text,
    docx_copy_paragraph, docx_delete, docx_insert_page_break,
)
from .run_tools import docx_insert_run, docx_update_run_text, docx_set_font
from .table_tools import (
    docx_insert_table, docx_get_table, docx_list_tables, docx_find_table,
    docx_get_cell, docx_insert_paragraph_to_cell, docx_insert_table_row,
    docx_insert_table_col, docx_fill_table, docx_copy_table, docx_get_table_structure,
)
from .table_rowcol_tools import (
    docx_insert_row_at, docx_insert_col_at, docx_delete_row, docx_delete_col,
)
from .format_tools import (
    docx_set_alignment, docx_set_properties, docx_format_copy,
    docx_set_margins, docx_extract_format_template, docx_apply_format_template,
)
from .advanced_tools import docx_replace_text, docx_batch_replace_text, docx_insert_image
from .composite_tools import (
    docx_insert_formatted_paragraph, docx_quick_edit, docx_get_structure_summary,
    docx_smart_fill_table, docx_format_range,
)
from .copy_tools import docx_get_element_source, docx_copy_elements_range

__all__ = [
    # Session/File
    "docx_save", "docx_get_context",
    # Content
    "docx_read_content", "docx_find_paragraphs", "docx_extract_template_structure",
    # Paragraph
    "docx_insert_paragraph", "docx_insert_heading", "docx_update_paragraph_text",
    "docx_copy_paragraph", "docx_delete", "docx_insert_page_break",
    # Run
    "docx_insert_run", "docx_update_run_text", "docx_set_font",
    # Table
    "docx_insert_table", "docx_get_table", "docx_list_tables", "docx_find_table",
    "docx_get_cell", "docx_insert_paragraph_to_cell", "docx_insert_table_row",
    "docx_insert_table_col", "docx_fill_table", "docx_copy_table", "docx_get_table_structure",
    "docx_insert_row_at", "docx_insert_col_at", "docx_delete_row", "docx_delete_col",
    # Format
    "docx_set_alignment", "docx_set_properties", "docx_format_copy",
    "docx_set_margins", "docx_extract_format_template", "docx_apply_format_template",
    # Advanced
    "docx_replace_text", "docx_batch_replace_text", "docx_insert_image",
    # Composite
    "docx_insert_formatted_paragraph", "docx_quick_edit", "docx_get_structure_summary",
    "docx_smart_fill_table", "docx_format_range",
    # Copy
    "docx_get_element_source", "docx_copy_elements_range",
]
