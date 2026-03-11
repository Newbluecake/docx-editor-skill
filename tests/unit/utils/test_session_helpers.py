"""Unit tests for session_helpers module."""

import pytest
from docx_editor_skill.utils.session_helpers import get_active_session
from docx_editor_skill.core.session import session_manager
from docx_editor_skill.core.global_state import global_state


class TestGetActiveSession:
    """Test suite for get_active_session() function."""

    def setup_method(self):
        """Setup test environment before each test."""
        # Clear global state
        global_state.clear()
        # Clean up any existing sessions
        for session_id in list(session_manager.sessions.keys()):
            session_manager.close_session(session_id)

    def teardown_method(self):
        """Cleanup after each test."""
        global_state.clear()
        for session_id in list(session_manager.sessions.keys()):
            session_manager.close_session(session_id)

    def test_no_active_session(self):
        """Test error when no active session exists."""
        session, error = get_active_session()

        assert session is None
        assert error is not None
        assert "NoActiveSession" in error
        assert "No active session" in error

    def test_session_not_found(self):
        """Test error when session_id exists but session is not in manager."""
        # Set a fake session ID that doesn't exist
        global_state.active_session_id = "fake_session_id"

        session, error = get_active_session()

        assert session is None
        assert error is not None
        assert "SessionNotFound" in error
        assert "fake_session_id" in error

    def test_valid_session(self):
        """Test successful retrieval of active session."""
        # Create a real session
        session_id = session_manager.create_session()
        global_state.active_session_id = session_id

        session, error = get_active_session()

        assert session is not None
        assert error is None
        assert session.session_id == session_id

    def test_session_with_file(self):
        """Test session retrieval with file path."""
        # Create session with file path
        import tempfile
        import os
        from docx import Document

        # Create a temporary docx file
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
            doc = Document()
            doc.save(tmp.name)
            tmp_path = tmp.name

        try:
            session_id = session_manager.create_session(file_path=tmp_path)
            global_state.active_session_id = session_id

            session, error = get_active_session()

            assert session is not None
            assert error is None
            assert session.file_path == tmp_path
        finally:
            # Cleanup
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)

    def test_multiple_calls(self):
        """Test that multiple calls return the same session."""
        session_id = session_manager.create_session()
        global_state.active_session_id = session_id

        session1, error1 = get_active_session()
        session2, error2 = get_active_session()

        assert session1 is not None
        assert session2 is not None
        assert error1 is None
        assert error2 is None
        assert session1.session_id == session2.session_id

    def test_session_switch(self):
        """Test switching between sessions."""
        # Create two sessions
        session_id1 = session_manager.create_session()
        session_id2 = session_manager.create_session()

        # Set first session as active
        global_state.active_session_id = session_id1
        session1, _ = get_active_session()
        assert session1.session_id == session_id1

        # Switch to second session
        global_state.active_session_id = session_id2
        session2, _ = get_active_session()
        assert session2.session_id == session_id2
        assert session2.session_id != session1.session_id
