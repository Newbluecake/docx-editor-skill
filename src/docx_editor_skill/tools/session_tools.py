"""Session management tools"""
import logging
from docx_editor_skill.core.response import create_markdown_response, create_error_response

logger = logging.getLogger(__name__)



def docx_get_current_session() -> str:
    """Get the current active session information.

    Returns information about the globally active session.

    Returns:
        str: Markdown-formatted response with session details
    """
    from docx_editor_skill.utils.session_helpers import get_active_session

    session, error = get_active_session()
    if error:
        return error

    return create_markdown_response(
        session=None,
        message="Current session retrieved successfully",
        operation="Get Current Session",
        show_context=False,
        session_id=session.session_id,
        file_path=session.file_path or "None",
        auto_save=session.auto_save,
        has_unsaved_changes=session.has_unsaved_changes()
    )


def docx_switch_session(session_id: str) -> str:
    """Switch to a different active session.

    Args:
        session_id: The session ID to switch to

    Returns:
        str: Markdown-formatted success message
    """
    from docx_editor_skill.core.session import session_manager
    from docx_editor_skill.core.global_state import global_state

    logger.info(f"docx_switch_session called: session_id={session_id}")

    session = session_manager.get_session(session_id)
    if not session:
        return create_error_response(
            message=f"Session {session_id} not found or expired",
            error_type="SessionNotFound"
        )

    global_state.active_session_id = session_id
    global_state.active_file = session.file_path

    logger.info(f"Switched to session {session_id}")
    return create_markdown_response(
        session=None,
        message=f"Switched to session {session_id}",
        operation="Switch Session",
        show_context=False,
        session_id=session_id,
        file_path=session.file_path or "None"
    )


def docx_save(
    file_path: str,
    backup: bool = False,
    backup_dir: str = "",
    backup_suffix: str = "",
) -> str:
    """
    Save the document to disk at the specified path.

    Persists all modifications made during the session to a .docx file.
    Supports saving to a new location or overwriting the original file.

    Typical Use Cases:
        - Save generated document to output location
        - Create a modified copy of a template
        - Checkpoint document state during long operations

    Args:
        file_path (str): Absolute or relative path where the file should be saved.
            Parent directory must exist. Use relative paths for portability.

    Returns:
        str: Markdown-formatted success message with the saved file path.

    Raises:
        ValueError: If session_id is invalid or session has expired.
        RuntimeError: If file cannot be saved (permission denied, disk full, etc.).
        FileNotFoundError: If parent directory does not exist.

    Examples:
        Save a new document:
        >>> # Session is automatically created by Launcher GUI or HTTP API
        >>> para_id = docx_insert_paragraph("Hello World", position="end:document_body")
        >>> result = docx_save("./output.docx")
        >>> print(result)
        'Document saved successfully to ./output.docx'

        Save modified template:
        >>> # Load template via Launcher GUI, session is auto-created
        >>> # ... make modifications ...
        >>> docx_save("./filled_template.docx")

    Notes:
        - If file exists, it will be overwritten without warning
        - In auto_save mode, this is called automatically after each modification
        - Live preview feature will refresh Word if file is open

    See Also:
        - docx_get_current_session: Get current session information
    """
    from docx_editor_skill.utils.session_helpers import get_active_session

    session, error = get_active_session()
    if error:
        return error

    logger.info(
        f"docx_save called: session_id={session.session_id}, file_path={file_path}, "
        f"backup={backup}, backup_dir={backup_dir}, backup_suffix={backup_suffix}"
    )

    # Security check: Ensure we are writing to an allowed location if needed.
    # For now, we assume the user (Claude) acts with the user's permissions.

    # LIVE PREVIEW: Prepare for save (release locks if file is open in Word)
    try:
        session.preview_controller.prepare_for_save(file_path)
    except Exception as e:
        # Log but don't block save
        logger.warning(f"Preview prepare failed: {e}")

    try:
        backup_path = session._save_with_optional_backup(
            file_path,
            backup=backup,
            backup_dir=backup_dir or None,
            backup_suffix=backup_suffix or None,
        )
        logger.info(f"docx_save success: {file_path}, backup={backup_path}")
    except Exception as e:
        logger.exception(f"docx_save failed: {e}")
        return create_error_response(
            message=f"Failed to save document: {str(e)}",
            error_type="SaveError"
        )

    # LIVE PREVIEW: Refresh (reload file in Word)
    try:
        session.preview_controller.refresh(file_path)
    except Exception as e:
         logger.warning(f"Preview refresh failed: {e}")

    # T-003: Mark session as saved (clear dirty flag)
    session.mark_saved()

    return create_markdown_response(
        session=None,
        message=f"Document saved successfully to {file_path}",
        operation="Save Document",
        show_context=False,
        file_path=file_path,
        backup_path=backup_path,
        backup=backup
    )

def docx_get_context() -> str:
    """
    Get the current context state of the session.

    Returns information about the session's current state, including the last created
    and accessed elements, file path, and auto-save status.

    Typical Use Cases:
        - Debug session state
        - Verify current context before operations
        - Check session configuration

    Args:
        None. Operates on the currently active session.

    Returns:
        str: Markdown-formatted session context information.

    Raises:
        ValueError: If session_id is invalid or session has expired.

    Examples:
        Check session context:
        >>> # Session is automatically created by Launcher GUI or HTTP API
        >>> para_id = docx_insert_paragraph("Text", position="end:document_body")
        >>> context = docx_get_context()
        >>> import json, re
        >>> match = re.search(r'\\*\\*Last Created ID\\*\\*:\\s*(\\w+)', context)
        >>> print(f"Last created: {match.group(1) if match else 'None'}")

    Notes:
        - Useful for debugging implicit context operations
        - Shows which element will be used for context-based operations
        - Auto-save status indicates if changes are automatically persisted

    See Also:
        - docx_get_current_session: Get current session information
        - docx_list_sessions: List all active sessions
    """
    from docx_editor_skill.utils.session_helpers import get_active_session

    session, error = get_active_session()
    if error:
        return error

    return create_markdown_response(
        session=None,
        message="Session context retrieved successfully",
        operation="Get Context",
        show_context=False,
        session_id=session.session_id,
        last_created_id=session.last_created_id or "None",
        last_accessed_id=session.last_accessed_id or "None",
        file_path=session.file_path or "None",
        auto_save=session.auto_save,
        backup_on_save=session.backup_on_save
    )


def docx_list_sessions() -> str:
    """
    List all active sessions with their metadata.

    Returns information about all currently active document editing sessions,
    including session IDs, file paths, and configuration settings.

    **No Session Required**: This tool operates at the server level and does
    not need a session_id. Use it to inspect active sessions for debugging
    or monitoring purposes.

    Typical Use Cases:
        - Debug session state and identify active sessions
        - Monitor server resource usage
        - Verify session creation before operations
        - Check which files are currently being edited

    Returns:
        str: Markdown-formatted response containing:
            - **Active Sessions**: Count of active sessions
            - **Sessions**: JSON array with session details (id, file_path, auto_save, etc.)

    Examples:
        List all active sessions:
        >>> result = docx_list_sessions()
        >>> # Response includes session count and details

        Check if specific session exists:
        >>> result = docx_list_sessions()
        >>> import json, re
        >>> match = re.search(r'\\*\\*Sessions\\*\\*:\\s*```json\\n(.+?)\\n```', result, re.DOTALL)
        >>> if match:
        ...     sessions = json.loads(match.group(1))
        ...     session_ids = [s['session_id'] for s in sessions]
        ...     print(f"Active sessions: {session_ids}")

    Notes:
        - Sessions auto-expire after 1 hour of inactivity
        - Each session maintains independent document state
        - Use docx_cleanup_sessions() to manually remove expired sessions

    See Also:
        - docx_get_current_session: Get current session information
        - docx_cleanup_sessions: Remove expired sessions
        - docx_get_context: Get detailed context for a specific session
    """
    from docx_editor_skill.core.session import session_manager

    sessions = session_manager.list_sessions()

    # Return Markdown format
    md_lines = ["# Active Sessions\n"]
    md_lines.append(f"**Count**: {len(sessions)}\n")

    if sessions:
        md_lines.append("## Session List\n")
        for idx, sess in enumerate(sessions, 1):
            md_lines.append(f"### Session {idx}")
            md_lines.append(f"- **ID**: `{sess.get('session_id', 'N/A')}`")
            md_lines.append(f"- **File**: {sess.get('file_path', 'N/A')}")
            md_lines.append(f"- **Auto Save**: {sess.get('auto_save', False)}")
            md_lines.append(f"- **Created**: {sess.get('created_at', 'N/A')}")
            md_lines.append("")
    else:
        md_lines.append("No active sessions.")

    return "\n".join(md_lines)


def docx_cleanup_sessions(max_idle_seconds: int = 0) -> str:
    """
    Clean up expired or idle sessions to free server resources.

    Removes sessions that have exceeded their idle timeout, releasing memory
    and clearing object registries. Can be called manually or runs automatically
    via the session manager's cleanup mechanism.

    **No Session Required**: This tool operates at the server level and does
    not need a session_id. Use it for server maintenance and resource management.

    Typical Use Cases:
        - Free memory after batch document processing
        - Clean up abandoned sessions in long-running servers
        - Force cleanup before server shutdown
        - Implement custom session timeout policies

    Args:
        max_idle_seconds (int, optional): Maximum idle time in seconds before
            a session is considered expired. If 0 or not provided, uses the
            default TTL (3600 seconds = 1 hour). Defaults to 0.

    Returns:
        str: Markdown-formatted response containing:
            - **Cleaned**: Number of sessions removed
            - **Max Idle Seconds**: Timeout threshold used

    Examples:
        Clean up with default timeout (1 hour):
        >>> result = docx_cleanup_sessions()
        >>> # Removes sessions idle for more than 1 hour

        Clean up with custom timeout (5 minutes):
        >>> result = docx_cleanup_sessions(max_idle_seconds=300)
        >>> # Removes sessions idle for more than 5 minutes

        Extract cleanup count:
        >>> result = docx_cleanup_sessions()
        >>> import re
        >>> match = re.search(r'\\*\\*Cleaned\\*\\*:\\s*(\\d+)', result)
        >>> count = int(match.group(1)) if match else 0
        >>> print(f"Cleaned up {count} sessions")

    Notes:
        - Closed sessions cannot be recovered - ensure work is saved first
        - Active sessions (recently accessed) are never cleaned up
        - Automatic cleanup runs periodically in the background
        - Use docx_list_sessions() to check active sessions before cleanup

    See Also:
        - docx_list_sessions: View active sessions before cleanup
        - docx_save: Save session before cleanup
    """
    from docx_editor_skill.core.session import session_manager

    cleaned = session_manager.cleanup_expired(max_idle_seconds=max_idle_seconds or None)
    return create_markdown_response(
        session=None,
        message=f"Cleaned up {cleaned} expired sessions",
        operation="Cleanup Sessions",
        show_context=False,
        cleaned=cleaned,
        max_idle_seconds=max_idle_seconds
    )
