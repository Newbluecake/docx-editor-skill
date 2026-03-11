"""Unit tests for docx_copy_paragraph functionality."""

import sys
import os
import json

# Add parent directory to path for helpers import
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from helpers import (
    extract_session_id,
    extract_element_id,
    extract_metadata_field,
    is_success,
    is_error
)
from tests.helpers.session_helpers import setup_active_session, teardown_active_session
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../../src')))

from docx_editor_skill.tools import (
    docx_insert_paragraph,
    docx_insert_run,
    docx_set_font,
    docx_set_alignment,
    docx_copy_paragraph,
    docx_read_content)


def test_copy_simple_paragraph():
    """Test copying a simple paragraph."""
    setup_active_session()
    # Create original paragraph
    para_response = docx_insert_paragraph("Original text", position="end:document_body")
    para_id = extract_element_id(para_response)

    # Copy it
    new_para_response = docx_copy_paragraph(para_id, position="end:document_body")
    new_para_id = extract_element_id(new_para_response)

    # Verify it's a different ID
    assert new_para_id != para_id
    assert new_para_id.startswith("para_")

    # Verify content
    content = docx_read_content()
    assert content.count("Original text") == 2

    teardown_active_session()

def test_copy_formatted_paragraph():
    """Test copying a paragraph with formatting."""
    setup_active_session()
    # Create formatted paragraph
    para_response = docx_insert_paragraph("", position="end:document_body")
    para_id = extract_element_id(para_response)
    run_response = docx_insert_run("Bold text", position=f"inside:{para_id}")
    run_id = extract_element_id(run_response)
    docx_set_font(run_id, bold=True, size=14, color_hex="FF0000")
    docx_set_alignment(para_id, "center")

    # Copy it
    new_para_response = docx_copy_paragraph(para_id, position="end:document_body")
    new_para_id = extract_element_id(new_para_response)

    # Verify different IDs
    assert new_para_id != para_id

    # Verify content duplicated
    content = docx_read_content()
    assert content.count("Bold text") == 2

    teardown_active_session()

def test_copy_paragraph_with_multiple_runs():
    """Test copying a paragraph with multiple runs."""
    setup_active_session()
    # Create paragraph with multiple runs
    para_response = docx_insert_paragraph("", position="end:document_body")
    para_id = extract_element_id(para_response)
    run1_response = docx_insert_run("Normal ", position=f"inside:{para_id}")
    run1_id = extract_element_id(run1_response)
    run2_response = docx_insert_run("Bold ", position=f"inside:{para_id}")
    run2_id = extract_element_id(run2_response)
    docx_set_font(run2_id, bold=True)
    run3_response = docx_insert_run("Italic", position=f"inside:{para_id}")
    run3_id = extract_element_id(run3_response)
    docx_set_font(run3_id, italic=True)

    # Copy it
    new_para_response = docx_copy_paragraph(para_id, position="end:document_body")
    new_para_id = extract_element_id(new_para_response)

    # Verify content
    content = docx_read_content()
    assert content.count("Normal Bold Italic") == 2

    teardown_active_session()

def test_copy_paragraph_invalid_session():
    """Test copying with no active session."""
    teardown_active_session()
    result = docx_copy_paragraph("para_123", position="end:document_body")
    assert is_error(result)
    assert "NoActiveSession" in result or "no active session" in result.lower()

def test_copy_paragraph_invalid_id():
    """Test copying with invalid paragraph ID."""
    setup_active_session()
    result = docx_copy_paragraph("para_invalid", position="end:document_body")
    try:
        data = json.loads(result)
        assert is_error(result)
        assert "not found" in data["message"].lower()
    except (json.JSONDecodeError, KeyError):
        assert "not found" in result.lower()

    teardown_active_session()

if __name__ == "__main__":
    test_copy_simple_paragraph()
    test_copy_formatted_paragraph()
    test_copy_paragraph_with_multiple_runs()
    test_copy_paragraph_invalid_session()
    test_copy_paragraph_invalid_id()
    print("All tests passed!")
