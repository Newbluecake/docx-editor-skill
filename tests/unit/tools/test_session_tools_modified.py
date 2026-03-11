"""Unit tests for modified session_tools (T-003)."""

import pytest
import tempfile
import os
from docx import Document
from docx_editor_skill.tools.session_tools import (
    docx_get_current_session,
    docx_switch_session,
    docx_save,
    docx_get_context
)
from docx_editor_skill.core.session import session_manager
from docx_editor_skill.core.global_state import global_state


class TestModifiedSessionTools:
    """Test suite for modified session tools without session_id parameter."""

    def setup_method(self):
        """Setup before each test."""
        global_state.clear()
        for sid in list(session_manager.sessions.keys()):
            session_manager.close_session(sid)

    def teardown_method(self):
        """Cleanup after each test."""
        global_state.clear()
        for sid in list(session_manager.sessions.keys()):
            session_manager.close_session(sid)

    def test_docx_get_current_session_no_active(self):
        """Test docx_get_current_session with no active session."""
        result = docx_get_current_session()
        assert "NoActiveSession" in result

    def test_docx_get_current_session_success(self):
        """Test docx_get_current_session with active session."""
        # Create and set active session
        session_id = session_manager.create_session()
        global_state.active_session_id = session_id

        result = docx_get_current_session()
        assert "Success" in result
        assert session_id in result

    def test_docx_switch_session_success(self):
        """Test docx_switch_session."""
        # Create two sessions
        sid1 = session_manager.create_session()
        sid2 = session_manager.create_session()

        # Switch to sid2
        result = docx_switch_session(sid2)
        assert "Success" in result
        assert global_state.active_session_id == sid2

    def test_docx_switch_session_not_found(self):
        """Test docx_switch_session with invalid session_id."""
        result = docx_switch_session("invalid_id")
        assert "SessionNotFound" in result

    def test_docx_save_no_active(self):
        """Test docx_save with no active session."""
        result = docx_save("/tmp/test.docx")
        assert "NoActiveSession" in result

    def test_docx_save_success(self):
        """Test docx_save with active session."""
        # Create temp file
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
            doc = Document()
            doc.save(tmp.name)
            tmp_path = tmp.name

        try:
            session_id = session_manager.create_session(tmp_path)
            global_state.active_session_id = session_id

            # Save to another location
            output_path = tmp_path.replace(".docx", "_output.docx")
            result = docx_save(output_path)

            assert "Success" in result
            assert os.path.exists(output_path)

            # Cleanup
            if os.path.exists(output_path):
                os.unlink(output_path)
        finally:
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)

    def test_docx_get_context_no_active(self):
        """Test docx_get_context with no active session."""
        result = docx_get_context()
        assert "NoActiveSession" in result

    def test_docx_get_context_success(self):
        """Test docx_get_context with active session."""
        session_id = session_manager.create_session()
        global_state.active_session_id = session_id

        result = docx_get_context()
        assert "Success" in result
        assert session_id in result
