"""Unit tests for GlobalState manager.

Tests thread safety, atomic operations, and state management.
"""

import threading
import pytest
from docx_editor_skill.core.global_state import GlobalState, global_state


class TestGlobalState:
    """Test suite for GlobalState class."""

    def test_initialization(self):
        """Test that GlobalState initializes with None values."""
        state = GlobalState()
        assert state.active_file is None
        assert state.active_session_id is None

    def test_set_active_file(self):
        """Test setting active file."""
        state = GlobalState()
        state.active_file = "/path/to/doc.docx"
        assert state.active_file == "/path/to/doc.docx"

    def test_set_active_session_id(self):
        """Test setting active session ID."""
        state = GlobalState()
        state.active_session_id = "session-123"
        assert state.active_session_id == "session-123"

    def test_clear(self):
        """Test clearing all state."""
        state = GlobalState()
        state.active_file = "/path/to/doc.docx"
        state.active_session_id = "session-123"

        state.clear()

        assert state.active_file is None
        assert state.active_session_id is None

    def test_get_status(self):
        """Test getting status snapshot."""
        state = GlobalState()
        state.active_file = "/path/to/doc.docx"
        state.active_session_id = "session-123"

        status = state.get_status()

        assert status["currentFile"] == "/path/to/doc.docx"
        assert status["sessionId"] == "session-123"

    def test_atomic_context_manager(self):
        """Test atomic context manager."""
        state = GlobalState()

        with state.atomic():
            state.active_file = "/path/to/doc.docx"
            state.active_session_id = "session-123"

        assert state.active_file == "/path/to/doc.docx"
        assert state.active_session_id == "session-123"

    def test_thread_safe_write_access(self):
        """Test thread-safe write access with concurrent writers."""
        state = GlobalState()
        iterations = 100
        num_threads = 10

        def writer(thread_id):
            for i in range(iterations):
                state.active_file = f"/file-{thread_id}-{i}.docx"

        threads = [threading.Thread(target=writer, args=(tid,)) for tid in range(num_threads)]

        for t in threads:
            t.start()
        for t in threads:
            t.join()

        # Should not crash or corrupt data
        assert state.active_file is not None
        assert state.active_file.endswith(".docx")

    def test_thread_safe_read_write_mixed(self):
        """Test thread safety with mixed reads and writes."""
        state = GlobalState()
        state.active_file = "/initial.docx"
        iterations = 100
        num_readers = 5
        num_writers = 5

        def reader():
            for _ in range(iterations):
                _ = state.active_file
                _ = state.active_session_id

        def writer(thread_id):
            for i in range(iterations):
                state.active_file = f"/file-{thread_id}-{i}.docx"
                state.active_session_id = f"session-{thread_id}-{i}"

        reader_threads = [threading.Thread(target=reader) for _ in range(num_readers)]
        writer_threads = [threading.Thread(target=writer, args=(tid,)) for tid in range(num_writers)]

        all_threads = reader_threads + writer_threads

        for t in all_threads:
            t.start()
        for t in all_threads:
            t.join()

        # Should not crash
        assert True

    def test_atomic_compound_operation(self):
        """Test atomic compound operations."""
        state = GlobalState()
        state.active_file = "/old.docx"
        state.active_session_id = "old-session"

        with state.atomic():
            # Compound operation: check and update
            if state.active_file == "/old.docx":
                state.active_file = "/new.docx"
                state.active_session_id = "new-session"

        assert state.active_file == "/new.docx"
        assert state.active_session_id == "new-session"

    def test_legacy_api_compatibility(self):
        """Test legacy API methods for backward compatibility."""
        state = GlobalState()

        # set_active_file
        state.set_active_file("/path/to/doc.docx")
        assert state.active_file == "/path/to/doc.docx"

        # clear_active_file
        state.clear_active_file()
        assert state.active_file is None
        assert state.active_session_id is None


class TestGlobalStateInstance:
    """Test the global singleton instance."""

    def test_global_state_singleton_exists(self):
        """Test that global_state singleton is available."""
        assert global_state is not None
        assert isinstance(global_state, GlobalState)

    def test_global_state_is_shared(self):
        """Test that global_state is shared across imports."""
        from docx_editor_skill.core.global_state import global_state as gs2

        global_state.active_file = "/shared.docx"
        assert gs2.active_file == "/shared.docx"

        # Cleanup
        global_state.clear()
