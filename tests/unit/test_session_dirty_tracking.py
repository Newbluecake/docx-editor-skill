"""Unit tests for Session dirty tracking (T-003).

Tests the dirty flag mechanism for tracking unsaved changes.
"""

import pytest
from unittest.mock import Mock, patch
from docx_editor_skill.core.session import Session
from docx import Document


@pytest.fixture
def mock_document():
    """Create a mock Document."""
    doc = Mock(spec=Document)
    doc.add_paragraph = Mock()
    doc.save = Mock()
    return doc


@pytest.fixture
def session(mock_document):
    """Create a test session."""
    return Session(
        session_id="test-session",
        document=mock_document
    )


class TestDirtyTracking:
    """Test suite for dirty tracking functionality."""

    def test_initial_state_clean(self, session):
        """Test that new session starts clean."""
        assert not session.has_unsaved_changes()
        assert session._is_dirty is False
        assert session._last_save_commit_index == -1

    def test_mark_dirty(self, session):
        """Test marking session as dirty."""
        session.mark_dirty()

        assert session.has_unsaved_changes()
        assert session._is_dirty is True

    def test_mark_saved(self, session):
        """Test marking session as saved."""
        # Make it dirty first
        session.mark_dirty()
        assert session.has_unsaved_changes()

        # Mark as saved
        session.mark_saved()
        assert not session.has_unsaved_changes()
        assert session._is_dirty is False

    def test_update_context_create_marks_dirty(self, session):
        """Test that update_context with action='create' marks dirty."""
        assert not session.has_unsaved_changes()

        session.update_context("para_123", action="create")

        assert session.has_unsaved_changes()
        assert session._is_dirty is True

    def test_update_context_update_marks_dirty(self, session):
        """Test that update_context with action='update' marks dirty."""
        assert not session.has_unsaved_changes()

        session.update_context("para_123", action="update")

        assert session.has_unsaved_changes()
        assert session._is_dirty is True

    def test_update_context_access_does_not_mark_dirty(self, session):
        """Test that update_context with action='access' does not mark dirty."""
        assert not session.has_unsaved_changes()

        session.update_context("para_123", action="access")

        assert not session.has_unsaved_changes()
        assert session._is_dirty is False

    def test_has_unsaved_changes_with_history_stack(self, session):
        """Test has_unsaved_changes considers history_stack."""
        from docx_editor_skill.core.commit import Commit

        # Add a commit to history
        session.history_stack.append(Commit.create(
            operation="test_operation",
            changes={},
            affected_elements=[]
        ))

        # Even without explicit dirty flag, should return True
        # because history_stack has grown since last save
        assert session.has_unsaved_changes()

    def test_mark_saved_updates_commit_index(self, session):
        """Test that mark_saved updates the last_save_commit_index."""
        from docx_editor_skill.core.commit import Commit

        # Add some commits
        session.history_stack.append(Commit.create(
            operation="test_operation_1",
            changes={},
            affected_elements=[]
        ))
        session.history_stack.append(Commit.create(
            operation="test_operation_2",
            changes={},
            affected_elements=[]
        ))

        # Mark as saved
        session.mark_saved()

        # Should record the current history length
        assert session._last_save_commit_index == 1  # len(history_stack) - 1

    def test_has_unsaved_after_new_commits(self, session):
        """Test that new commits after save are detected as unsaved."""
        from docx_editor_skill.core.commit import Commit

        # Add a commit and save
        session.history_stack.append(Commit.create(
            operation="test_operation_1",
            changes={},
            affected_elements=[]
        ))
        session.mark_saved()
        assert not session.has_unsaved_changes()

        # Add another commit
        session.history_stack.append(Commit.create(
            operation="test_operation_2",
            changes={},
            affected_elements=[]
        ))

        # Should now have unsaved changes
        assert session.has_unsaved_changes()

    def test_thread_safety_mark_dirty(self, session):
        """Test thread safety of mark_dirty."""
        import threading

        def mark_dirty_multiple():
            for _ in range(100):
                session.mark_dirty()

        threads = [threading.Thread(target=mark_dirty_multiple) for _ in range(10)]

        for t in threads:
            t.start()
        for t in threads:
            t.join()

        # Should be dirty after all threads complete
        assert session.has_unsaved_changes()

    def test_thread_safety_mark_saved(self, session):
        """Test thread safety of mark_saved."""
        import threading

        session.mark_dirty()

        def mark_saved_multiple():
            for _ in range(100):
                session.mark_saved()

        threads = [threading.Thread(target=mark_saved_multiple) for _ in range(10)]

        for t in threads:
            t.start()
        for t in threads:
            t.join()

        # Should be clean after all threads complete
        assert not session.has_unsaved_changes()

    def test_auto_save_clears_dirty_flag(self, session, mock_document):
        """Test that auto-save with mark_saved clears the dirty flag."""
        # Enable auto-save
        session.auto_save = True
        session.file_path = "/tmp/test.docx"

        # Mock the save method
        with patch.object(session, '_save_with_optional_backup') as mock_save:
            # Mark dirty and trigger auto-save via update_context
            session.update_context("para_123", action="create")

            # Auto-save should have been triggered
            mock_save.assert_called_once()

            # Dirty flag should be cleared by mark_saved in update_context
            # Note: The actual implementation calls mark_saved() after successful auto-save
            assert not session.has_unsaved_changes()
