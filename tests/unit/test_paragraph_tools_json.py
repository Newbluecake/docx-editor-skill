"""Unit tests for refactored paragraph tools with JSON responses."""

import json
import pytest
from docx_editor_skill.tools.paragraph_tools import (
    docx_insert_paragraph,
    docx_insert_heading,
    docx_update_paragraph_text,
    docx_copy_paragraph,
    docx_delete,
    docx_insert_page_break
)
from tests.helpers.session_helpers import setup_active_session, teardown_active_session
from tests.helpers import (
    extract_session_id,
    extract_element_id,
    extract_metadata_field,
    extract_error_message,
    is_success,
    is_error
)


def test_add_paragraph_returns_json(active_session):
    """Test that docx_insert_paragraph returns valid JSON."""
    result = docx_insert_paragraph("Test paragraph", position="end:document_body")

    # Parse JSON
    assert is_success(result)
    assert extract_metadata_field(result, "element_id") is not None
    assert extract_metadata_field(result, "element_id").startswith("para_")


def test_add_paragraph_with_style(active_session):
    """Test adding paragraph with style returns proper JSON."""
    result = docx_insert_paragraph("Bullet item", position="end:document_body", style="List Bullet")
    assert is_success(result)
    assert extract_metadata_field(result, "element_id") is not None


def test_add_paragraph_no_active_session():
    """Test that calling without active session returns NoActiveSession error."""
    # No active session - should return error
    result = docx_insert_paragraph("Test", position="end:document_body")
    assert is_error(result)
    assert extract_metadata_field(result, "error_type") == "NoActiveSession"


def test_add_heading_returns_json(active_session):
    """Test that docx_insert_heading returns valid JSON."""
    result = docx_insert_heading("Chapter 1", position="end:document_body", level=1)
    assert is_success(result)
    assert extract_metadata_field(result, "element_id") is not None
    assert extract_metadata_field(result, "element_id").startswith("para_")


def test_add_heading_with_different_levels(active_session):
    """Test adding headings with different levels."""
    # Title (level 0)
    result = docx_insert_heading("Title", position="end:document_body", level=0)
    assert is_success(result)

    # Heading 2
    result = docx_insert_heading("Section", position="end:document_body", level=2)
    assert is_success(result)


def test_update_paragraph_text_returns_json(active_session):
    """Test that docx_update_paragraph_text returns valid JSON."""
    # Create paragraph
    result = docx_insert_paragraph("Old text", position="end:document_body")
    para_id = extract_element_id(result)

    # Update it
    result = docx_update_paragraph_text(para_id, "New text")

    assert is_success(result)
    assert extract_metadata_field(result, "element_id") == para_id
    assert extract_metadata_field(result, "changed_fields") is not None


def test_update_paragraph_invalid_id(active_session):
    """Test updating non-existent paragraph returns error."""
    result = docx_update_paragraph_text("para_invalid", "New text")
    assert is_error(result)
    assert extract_metadata_field(result, "error_type") == "ElementNotFound"


def test_copy_paragraph_returns_json(active_session):
    """Test that docx_copy_paragraph returns valid JSON."""
    # Create original paragraph
    result = docx_insert_paragraph("Original text", position="end:document_body")
    original_id = extract_element_id(result)

    # Copy it
    result = docx_copy_paragraph(original_id, position="end:document_body")

    assert is_success(result)
    assert extract_metadata_field(result, "element_id") is not None
    assert extract_element_id(result) != original_id


def test_delete_paragraph_returns_json(active_session):
    """Test that docx_delete returns valid JSON."""
    # Create paragraph
    result = docx_insert_paragraph("To be deleted", position="end:document_body")
    para_id = extract_element_id(result)

    # Delete it
    result = docx_delete(para_id)

    assert is_success(result)
    assert extract_metadata_field(result, "deleted_id") == para_id


def test_delete_without_element_id_uses_context(active_session):
    """Test that delete without element_id uses context."""
    # Create paragraph (sets context)
    result = docx_insert_paragraph("Context paragraph", position="end:document_body")
    para_id = extract_element_id(result)

    # Delete using context
    result = docx_delete()

    assert is_success(result)
    assert extract_metadata_field(result, "deleted_id") == para_id


def test_delete_no_context_returns_error(active_session):
    """Test that delete without context returns error."""
    # Try to delete without any context
    result = docx_delete()

    assert is_error(result)
    assert extract_metadata_field(result, "error_type") == "NoContext"


def test_add_page_break_returns_json(active_session):
    """Test that docx_insert_page_break returns valid JSON."""
    result = docx_insert_page_break(position="end:document_body")
    assert is_success(result)


def test_cursor_context_in_response(active_session):
    """Test that cursor context is included in responses."""
    # Add first paragraph
    docx_insert_paragraph("First", position="end:document_body")

    # Add second paragraph
    result = docx_insert_paragraph("Second", position="end:document_body")

    # Check that response includes document context (cursor is shown in context)
    assert is_success(result)
    assert "Document Context" in result or "CURSOR" in result


def test_json_response_structure(active_session):
    """Test that all responses follow the standard structure."""
    # Test various operations
    operations = [
        docx_insert_paragraph("Test", position="end:document_body"),
        docx_insert_heading("Heading", position="end:document_body", level=1),
        docx_insert_page_break(position="end:document_body")
    ]

    for result in operations:
        # All responses must be successful and have element_id
        assert is_success(result)
        assert extract_element_id(result) is not None


def test_error_response_structure():
    """Test that error responses follow the standard structure."""
    # Test without active session
    result = docx_insert_paragraph("Test", position="end:document_body")

    assert is_error(result)
    assert extract_error_message(result) is not None
    assert extract_metadata_field(result, "error_type") == "NoActiveSession"


def test_add_paragraph_to_parent(active_session):
    """Test adding paragraph to a parent container (cell)."""
    from docx_editor_skill.tools.table_tools import docx_insert_table, docx_get_cell

    # Create table
    table_result = docx_insert_table(2, 2, position="end:document_body")
    table_id = extract_element_id(table_result)

    # Get cell
    cell_result = docx_get_cell(table_id, 0, 0)
    cell_id = extract_element_id(cell_result)

    # Add paragraph to cell
    result = docx_insert_paragraph("Cell content", position=f"inside:{cell_id}")

    assert is_success(result)
    assert extract_metadata_field(result, "element_id") is not None
