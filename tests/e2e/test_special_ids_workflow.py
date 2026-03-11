"""
E2E tests for special position IDs feature.

Tests the following scenarios:
1. Simplified consecutive insertion using last_insert
2. Using cursor for positioning
3. Format copy workflow with last_insert + last_update
4. Error recovery when using special IDs before initialization
"""

import pytest
from tests.helpers import (
    extract_session_id,
    extract_element_id,
    extract_metadata_field,
    is_success,
    is_error,
    extract_error_message,
    create_session_with_file,
)
from tests.helpers.session_helpers import setup_active_session, teardown_active_session
from docx_editor_skill.tools.session_tools import docx_save
from docx_editor_skill.tools.paragraph_tools import docx_insert_paragraph
from docx_editor_skill.tools.run_tools import docx_insert_run, docx_set_font
from docx_editor_skill.tools.format_tools import docx_format_copy
from docx_editor_skill.tools.table_tools import docx_insert_table
from docx_editor_skill.core.session import session_manager
import tempfile
import os


def test_consecutive_insertion_with_last_insert():
    """Test simplified consecutive insertion using last_insert special ID."""
    # Create session
    session_id = setup_active_session()
    # Insert first paragraph
    p1_resp = docx_insert_paragraph("First paragraph", position="end:document_body")
    assert is_success(p1_resp)
    p1_id = extract_element_id(p1_resp)

    # Insert second paragraph using last_insert (no need to extract p1_id)
    p2_resp = docx_insert_paragraph("Second paragraph", position="after:last_insert")
    assert is_success(p2_resp)
    p2_id = extract_element_id(p2_resp)

    # Insert third paragraph using last_insert again
    p3_resp = docx_insert_paragraph("Third paragraph", position="after:last_insert")
    assert is_success(p3_resp)
    p3_id = extract_element_id(p3_resp)

    # Insert table using last_insert
    table_resp = docx_insert_table(2, 2, position="after:last_insert")
    assert is_success(table_resp)
    table_id = extract_element_id(table_resp)

    # Verify order in document
    session = session_manager.get_session(session_id)
    doc = session.document
    body_elements = doc._body._element.getchildren()

    def get_text_or_type(el):
        if el.tag.endswith('p'):
            return "P: " + "".join([t.text for t in el.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")])
        elif el.tag.endswith('tbl'):
            return "Table"
        return "Unknown"

    content = [get_text_or_type(el) for el in body_elements if el.tag.endswith(('p', 'tbl'))]

    assert "P: First paragraph" in content[0]
    assert "P: Second paragraph" in content[1]
    assert "P: Third paragraph" in content[2]
    assert "Table" in content[3]

    # Cleanup
    teardown_active_session()


def test_cursor_positioning():
    """Test using cursor for positioning with special IDs."""
    session_id = setup_active_session()
    # Insert initial paragraphs
    p1_resp = docx_insert_paragraph("Paragraph 1", position="end:document_body")
    assert is_success(p1_resp)

    p2_resp = docx_insert_paragraph("Paragraph 2", position="after:last_insert")
    assert is_success(p2_resp)

    p3_resp = docx_insert_paragraph("Paragraph 3", position="after:last_insert")
    assert is_success(p3_resp)

    # Now insert using cursor (which should be at p3)
    p4_resp = docx_insert_paragraph("Paragraph 4", position="after:cursor")
    assert is_success(p4_resp)

    # Verify order
    session = session_manager.get_session(session_id)
    doc = session.document
    paragraphs = [p.text for p in doc.paragraphs]

    assert paragraphs[0] == "Paragraph 1"
    assert paragraphs[1] == "Paragraph 2"
    assert paragraphs[2] == "Paragraph 3"
    assert paragraphs[3] == "Paragraph 4"

    # Cleanup
    teardown_active_session()


def test_format_copy_with_special_ids():
    """Test format copy workflow using last_insert and last_update."""
    session_id = setup_active_session()
    # Create source paragraph with formatted run
    p1_resp = docx_insert_paragraph("", position="end:document_body")
    assert is_success(p1_resp)
    p1_id = extract_element_id(p1_resp)

    # Add formatted run to source
    run1_resp = docx_insert_run("Bold Red Text", position=f"inside:{p1_id}")
    assert is_success(run1_resp)
    run1_id = extract_element_id(run1_resp)

    # Format the run (this updates last_update to run1_id)
    font_resp = docx_set_font(run1_id, bold=True, color_hex="FF0000", size=16)
    assert is_success(font_resp)

    # Create target paragraph (use after:p1_id instead of after:last_insert which is run1)
    p2_resp = docx_insert_paragraph("", position=f"after:{p1_id}")
    assert is_success(p2_resp)
    p2_id = extract_element_id(p2_resp)

    # Add run to target (this updates last_insert to run2_id)
    run2_resp = docx_insert_run("Normal Text", position=f"inside:{p2_id}")
    assert is_success(run2_resp)
    run2_id = extract_element_id(run2_resp)

    # Copy format from run1 (use explicit ID) to run2 (use last_insert)
    format_resp = docx_format_copy(source_id=run1_id, target_id="last_insert")
    assert is_success(format_resp)

    # Verify formatting was copied
    session = session_manager.get_session(session_id)
    run2 = session.get_object(run2_id)

    assert run2.bold is True
    assert run2.font.size.pt == 16
    # Color check (RGB)
    assert run2.font.color.rgb is not None

    # Cleanup
    teardown_active_session()


def test_error_recovery_uninitialized_special_ids():
    """Test error handling when using special IDs before initialization."""
    session_id = setup_active_session()
    # Try to use last_insert before any insertion
    p1_resp = docx_insert_paragraph("Test", position="after:last_insert")
    assert is_error(p1_resp)
    error_msg = extract_error_message(p1_resp)
    assert "last_insert" in error_msg.lower() or "not initialized" in error_msg.lower()

    # Try to use last_update before any update
    p2_resp = docx_insert_paragraph("Test", position="after:last_update")
    assert is_error(p2_resp)
    error_msg = extract_error_message(p2_resp)
    assert "last_update" in error_msg.lower() or "not initialized" in error_msg.lower()

    # Now do a valid insertion
    p3_resp = docx_insert_paragraph("Valid", position="end:document_body")
    assert is_success(p3_resp)

    # Now last_insert should work
    p4_resp = docx_insert_paragraph("After valid", position="after:last_insert")
    assert is_success(p4_resp)

    # Cleanup
    teardown_active_session()


def test_special_ids_with_file_save():
    """Test that special IDs work correctly with file save operations."""
    session_id = setup_active_session()
    # Create content using special IDs
    p1_resp = docx_insert_paragraph("Title", position="end:document_body")
    assert is_success(p1_resp)

    p2_resp = docx_insert_paragraph("Content 1", position="after:last_insert")
    assert is_success(p2_resp)

    p3_resp = docx_insert_paragraph("Content 2", position="after:last_insert")
    assert is_success(p3_resp)

    table_resp = docx_insert_table(2, 2, position="after:last_insert")
    assert is_success(table_resp)

    # Save to temporary file
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp_path = tmp.name

    try:
        save_resp = docx_save(tmp_path)
        assert is_success(save_resp)

        # Verify file exists and is valid
        assert os.path.exists(tmp_path)
        assert os.path.getsize(tmp_path) > 0

        # Load the file and verify content
        session2_resp = create_session_with_file(tmp_path)
        session2_id = extract_session_id(session2_resp)

        session2 = session_manager.get_session(session2_id)
        doc = session2.document

        paragraphs = [p.text for p in doc.paragraphs]
        assert "Title" in paragraphs[0]
        assert "Content 1" in paragraphs[1]
        assert "Content 2" in paragraphs[2]

        assert len(doc.tables) == 1

        # Cleanup - session will auto-expire, no need to close explicitly

    finally:
        # Clean up temp file
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)

    # Cleanup
    teardown_active_session()


def test_mixed_special_and_explicit_ids():
    """Test mixing special IDs with explicit element IDs."""
    session_id = setup_active_session()
    # Insert with explicit position
    p1_resp = docx_insert_paragraph("Paragraph 1", position="end:document_body")
    assert is_success(p1_resp)
    p1_id = extract_element_id(p1_resp)

    # Insert using last_insert
    p2_resp = docx_insert_paragraph("Paragraph 2", position="after:last_insert")
    assert is_success(p2_resp)
    p2_id = extract_element_id(p2_resp)

    # Insert using explicit ID (p1_id)
    p3_resp = docx_insert_paragraph("Paragraph 1.5", position=f"after:{p1_id}")
    assert is_success(p3_resp)
    p3_id = extract_element_id(p3_resp)

    # Insert using last_insert (should be after p3)
    p4_resp = docx_insert_paragraph("Paragraph 1.75", position="after:last_insert")
    assert is_success(p4_resp)

    # Verify order: P1, P1.5, P1.75, P2
    session = session_manager.get_session(session_id)
    doc = session.document
    paragraphs = [p.text for p in doc.paragraphs]

    assert paragraphs[0] == "Paragraph 1"
    assert paragraphs[1] == "Paragraph 1.5"
    assert paragraphs[2] == "Paragraph 1.75"
    assert paragraphs[3] == "Paragraph 2"

    # Cleanup
    teardown_active_session()

