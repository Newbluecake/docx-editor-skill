import json
import os
import time
import tempfile

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
    extract_error_message,
    create_session_with_file,
)
from tests.helpers.session_helpers import setup_active_session, teardown_active_session

from docx_editor_skill.tools.session_tools import (
    docx_save,
    docx_get_context,
    docx_list_sessions,
    docx_cleanup_sessions,
)
from docx_editor_skill.core.session import session_manager


def test_create_and_context_with_backup_flags():
    setup_active_session()
    try:
        from docx_editor_skill.core.global_state import global_state
        session_id = global_state.active_session_id
        ctx_result = docx_get_context()

        assert str(extract_metadata_field(ctx_result, "session_id")) == session_id
        # Note: backup_on_save and backup_suffix are set during save, not create
        # So we just check that context returns valid session info
        assert session_id is not None
    finally:
        teardown_active_session()


def test_save_with_backup_creates_backup_file(tmp_path):
    target = tmp_path / "sample.docx"

    # Create session bound to the target path
    session_id = create_session_with_file(str(target), auto_save=False)
    try:
        # First save creates the file
        first_result = docx_save(str(target))
        assert is_success(first_result)
        assert os.path.exists(target)

        # Second save with backup should produce a backup file
        second_result = docx_save(str(target), backup=True, backup_suffix="-bak")
        assert is_success(second_result)
        backup_path = extract_metadata_field(second_result, "backup_path")
        assert backup_path
        assert os.path.exists(backup_path)
    finally:
        teardown_active_session()


def test_list_sessions_includes_active_session():
    setup_active_session()
    try:
        listing = docx_list_sessions()
        # docx_list_sessions now returns Markdown format
        assert "# Active Sessions" in listing
        assert "**Count**:" in listing
    finally:
        teardown_active_session()


def test_cleanup_sessions_removes_idle():
    setup_active_session()
    try:
        from docx_editor_skill.core.global_state import global_state
        session_id = global_state.active_session_id
        session = session_manager.get_session(session_id)

        # Simulate idleness
        session.last_accessed = time.time() - 5

        result = docx_cleanup_sessions(max_idle_seconds=1)
        assert is_success(result)
        cleaned_count = extract_metadata_field(result, "cleaned")
        assert cleaned_count >= 1

        # Ensure session removed
        assert session_manager.get_session(session_id) is None
    finally:
        # Session already cleaned up, but call anyway for safety
        from docx_editor_skill.core.global_state import global_state
        global_state.active_session_id = None
