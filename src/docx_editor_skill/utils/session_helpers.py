"""Session helper utilities for accessing the global active session.

This module provides helper functions to simplify session access in tools.
Instead of requiring session_id as a parameter, tools can use get_active_session()
to automatically retrieve the currently active session from global state.

Example:
    >>> from docx_editor_skill.utils.session_helpers import get_active_session
    >>>
    >>> def docx_insert_paragraph(text: str, position: str) -> str:
    ...     session, error = get_active_session()
    ...     if error:
    ...         return error
    ...     # Use session...
"""

import logging
from typing import Optional, Tuple, Any

logger = logging.getLogger(__name__)


def get_active_session() -> Tuple[Optional[Any], Optional[str]]:
    """Get the active session from global state.

    This function retrieves the currently active session from the global state.
    It handles error cases where no session is active or the session has expired.

    Returns:
        tuple: (session, error_response)
            - If successful: (Session object, None)
            - If failed: (None, error response string in Markdown format)

    Error Cases:
        - NoActiveSession: No session_id in global_state (user hasn't switched files)
        - SessionNotFound: Session exists in global_state but not in session_manager
          (session expired or was closed)
    """
    from docx_editor_skill.core.session import session_manager
    from docx_editor_skill.core.global_state import global_state
    from docx_editor_skill.core.response import create_error_response

    # Get active session ID from global state
    session_id = global_state.active_session_id

    # Check if session ID exists
    if not session_id:
        logger.debug("No active session in global state")
        return None, create_error_response(
            message="No active session. Please open a file first.",
            error_type="NoActiveSession"
        )

    # Get session from session manager
    session = session_manager.get_session(session_id)

    # Check if session exists and is valid
    if not session:
        logger.warning(f"Active session {session_id} not found or expired")
        return None, create_error_response(
            message=f"Active session {session_id} not found or expired. The session may have timed out.",
            error_type="SessionNotFound"
        )

    logger.debug(f"Retrieved active session: {session_id}")
    return session, None
