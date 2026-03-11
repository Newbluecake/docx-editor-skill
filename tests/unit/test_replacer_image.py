import pytest
import os
import json
from unittest.mock import MagicMock, patch
from docx_editor_skill.core.replacer import replace_text_in_paragraph
from docx_editor_skill.core.session import session_manager
from docx_editor_skill.tools import docx_replace_text, docx_insert_image, docx_delete, docx_get_context

# Add parent directory to path for helpers import
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from helpers import (
    extract_session_id,
    extract_element_id,
    extract_metadata_field,
    is_success,
    is_error,
    extract_error_message
)
from tests.helpers.session_helpers import setup_active_session, teardown_active_session


def test_replace_text_single_run():
    # Setup
    mock_para = MagicMock()
    mock_run = MagicMock()
    mock_run.text = "Hello {{name}}!"
    mock_para.runs = [mock_run]
    mock_para.text = "Hello {{name}}!"

    # Action
    replace_text_in_paragraph(mock_para, "{{name}}", "Claude")

    # Assert
    assert mock_run.text == "Hello Claude!"

def test_replace_text_split_runs():
    # Setup: "{{name}}" split across runs
    mock_para = MagicMock()
    run1 = MagicMock()
    run1.text = "Hello {{"
    run2 = MagicMock()
    run2.text = "name}}!"

    mock_para.runs = [run1, run2]
    mock_para.text = "Hello {{name}}!"

    # Action
    replace_text_in_paragraph(mock_para, "{{name}}", "Claude")

    # Assert cross-run replacement preserved runs instead of clearing
    assert run1.text == "Hello Claude!"
    assert run2.text == ""

def test_docx_insert_image():
    with patch("os.path.exists", return_value=True):
        setup_active_session()
        try:
            from docx_editor_skill.core.global_state import global_state
            sid = global_state.active_session_id
            # Mock add_picture on run
            with patch("docx.text.run.Run.add_picture") as mock_add_pic:
                 # Insert
                 r_response = docx_insert_image("/tmp/image.png", width=2.0, position="end:document_body")
                 r_id = extract_element_id(r_response)

                 assert r_id.startswith("run_") or r_id.startswith("para_") # Our implementation returns para_id by default if parent not specified

                 # Verify session update
                 session = session_manager.get_session(sid)
                 clean_id = r_id.strip().split()[0]
                 assert session.last_created_id == clean_id
        finally:
            teardown_active_session()

def test_context_and_delete():
    setup_active_session()
    try:
        from docx_editor_skill.core.global_state import global_state
        sid = global_state.active_session_id
        session = session_manager.get_session(sid)

        # Check context
        ctx = docx_get_context()
        assert extract_metadata_field(ctx, "last_created_id") is not None

        # Delete something
        # Mock an object with _element for deletion logic
        mock_obj = MagicMock()
        mock_xml = MagicMock()
        mock_parent = MagicMock()
        mock_xml.getparent.return_value = mock_parent
        mock_obj._element = mock_xml

        session.object_registry["obj_1"] = mock_obj

        docx_delete("obj_1")

        # Verify removal
        mock_parent.remove.assert_called_with(mock_xml)
        assert "obj_1" not in session.object_registry
    finally:
        teardown_active_session()
