"""Session management helpers for testing.

This module provides helper functions for setting up and tearing down
test sessions. Uses CLISession with deterministic IDs for reproducible tests.
"""

from typing import Optional
from docx_editor_skill.core.session import session_manager
from docx_editor_skill.core.global_state import global_state


def setup_active_session(file_path: Optional[str] = None) -> str:
    """Setup a global active session for testing.

    Creates a session and sets it as the active session in global_state.

    Args:
        file_path: Optional file path to load. If None, creates an empty document.

    Returns:
        str: Created session_id
    """
    session_id = session_manager.create_session(file_path)
    global_state.active_session_id = session_id
    if file_path:
        global_state.active_file = file_path
    return session_id


def teardown_active_session():
    """Teardown the global active session.

    Closes the active session and clears global_state.
    Should be called in test cleanup to prevent state leakage.
    """
    if global_state.active_session_id:
        session_manager.close_session(global_state.active_session_id)
    global_state.clear()


def create_session_with_file(file_path: str, **kwargs) -> str:
    """Create a session with a specific file (legacy compatibility).

    Args:
        file_path: Path to the file to load
        **kwargs: Additional arguments (auto_save, backup_on_save, etc.)

    Returns:
        str: Created session_id
    """
    session_id = session_manager.create_session(
        file_path=file_path,
        auto_save=kwargs.get('auto_save', False),
        backup_on_save=kwargs.get('backup_on_save', False),
        backup_dir=kwargs.get('backup_dir'),
        backup_suffix=kwargs.get('backup_suffix')
    )
    global_state.active_session_id = session_id
    global_state.active_file = file_path
    return session_id


def clear_active_file():
    """Clear the global active file (useful for test cleanup)."""
    global_state.active_file = None
    global_state.active_session_id = None
