"""Unit tests for text update functionality."""

import sys
import os
import json
import re
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../../src')))

from helpers import extract_session_id, extract_element_id
from tests.helpers.session_helpers import setup_active_session, teardown_active_session
from docx_editor_skill.tools import (
    docx_insert_paragraph,
    docx_insert_run,
    docx_set_font,
    docx_update_paragraph_text,
    docx_update_run_text,
    docx_read_content)


# extract_element_id removed


# extract_session_id removed

def test_update_paragraph_text():
    """Test updating paragraph text."""
    setup_active_session()
    # Create paragraph
    para_response = docx_insert_paragraph("Original text", position="end:document_body")
    para_id = extract_element_id(para_response)

    # Update it
    result = docx_update_paragraph_text(para_id, "Updated text")
    # Result can be JSON or plain string
    result_str = result if isinstance(result, str) else str(result)
    assert "updated" in result_str.lower() or "success" in result_str.lower()

    # Verify content
    content = docx_read_content()
    assert "Updated text" in content
    assert "Original text" not in content

    teardown_active_session()

def test_update_run_text():
    """Test updating run text while preserving formatting."""
    setup_active_session()
    # Create formatted run
    para_response = docx_insert_paragraph("", position="end:document_body")
    para_id = extract_element_id(para_response)
    run_response = docx_insert_run("Original", position=f"inside:{para_id}")
    run_id = extract_element_id(run_response)
    docx_set_font(run_id, bold=True, size=16)

    # Update run text
    result = docx_update_run_text(run_id, "Updated")
    result_str = result if isinstance(result, str) else str(result)
    assert "updated" in result_str.lower() or "success" in result_str.lower()

    # Verify content
    content = docx_read_content()
    assert "Updated" in content
    assert "Original" not in content

    teardown_active_session()

def test_update_paragraph_with_multiple_runs():
    """Test updating paragraph replaces all runs."""
    setup_active_session()
    # Create paragraph with multiple runs
    para_response = docx_insert_paragraph("", position="end:document_body")
    para_id = extract_element_id(para_response)
    docx_insert_run("Part 1 ", position=f"inside:{para_id}")
    docx_insert_run("Part 2", position=f"inside:{para_id}")

    # Update paragraph (should replace all runs)
    docx_update_paragraph_text(para_id, "New single text")

    # Verify content
    content = docx_read_content()
    assert "New single text" in content
    assert "Part 1" not in content
    assert "Part 2" not in content

    teardown_active_session()

def test_update_paragraph_invalid_session():
    """Test updating paragraph when no active session exists."""
    # Ensure no active session
    teardown_active_session()

    result = docx_update_paragraph_text("para_123", "text")
    # Should return NoActiveSession error
    assert "NoActiveSession" in result or "no active session" in result.lower()

def test_update_paragraph_invalid_id():
    """Test updating with invalid paragraph ID."""
    setup_active_session()
    result = docx_update_paragraph_text("para_invalid", "text")
    # Should return error JSON response
    try:
        data = json.loads(result)
        assert data["status"] == "error"
        assert "not found" in data["message"].lower()
    except (json.JSONDecodeError, KeyError):
        assert "not found" in result.lower()

    teardown_active_session()

def test_update_run_invalid_session():
    """Test updating run when no active session exists."""
    # Ensure no active session
    teardown_active_session()

    result = docx_update_run_text("run_123", "text")
    # Should return NoActiveSession error
    assert "NoActiveSession" in result or "no active session" in result.lower()

def test_update_run_invalid_id():
    """Test updating run with invalid run ID."""
    setup_active_session()
    result = docx_update_run_text("run_invalid", "text")
    # Should return error JSON response
    try:
        data = json.loads(result)
        assert data["status"] == "error"
        assert "not found" in data["message"].lower()
    except (json.JSONDecodeError, KeyError):
        assert "not found" in result.lower()

    teardown_active_session()

if __name__ == "__main__":
    test_update_paragraph_text()
    test_update_run_text()
    test_update_paragraph_with_multiple_runs()
    test_update_paragraph_invalid_session()
    test_update_paragraph_invalid_id()
    test_update_run_invalid_session()
    test_update_run_invalid_id()
    print("All tests passed!")
