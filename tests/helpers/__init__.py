"""Test helper utilities for docx-editor-skill tests."""

from .markdown_extractors import (
    extract_session_id,
    extract_element_id,
    extract_status,
    extract_metadata_field,
    extract_all_metadata,
    is_success,
    is_error,
    extract_error_message,
    has_metadata_field,
    extract_json_from_markdown,
    extract_find_paragraphs_results,
    extract_template_structure,
    extract_element_ids_list,
)
from .session_helpers import (
    create_session_with_file,
    clear_active_file,
    setup_active_session,
    teardown_active_session,
)

__all__ = [
    'extract_session_id',
    'extract_element_id',
    'extract_status',
    'extract_metadata_field',
    'extract_all_metadata',
    'is_success',
    'is_error',
    'extract_error_message',
    'has_metadata_field',
    'extract_json_from_markdown',
    'extract_find_paragraphs_results',
    'extract_template_structure',
    'extract_element_ids_list',
    'create_session_with_file',
    'clear_active_file',
    'setup_active_session',
    'teardown_active_session',
]
