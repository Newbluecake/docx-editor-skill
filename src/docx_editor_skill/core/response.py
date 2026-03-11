"""Response formatting utilities for MCP tools.

This module provides Markdown-formatted responses with ASCII visualization
for all tools, replacing the previous JSON-based format.
"""

import logging
import json
from typing import Optional, Any

logger = logging.getLogger(__name__)


# Error suggestions for common error types
ERROR_SUGGESTIONS = {
    "SessionNotFound": "Check that the session ID is correct and the session has not expired.",
    "ElementNotFound": "Verify that the element ID exists and has not been deleted.",
    "InvalidElementType": "Ensure you are using the correct tool for the element type.",
    "ValidationError": "Check that all parameters meet the required format and constraints.",
    "FileNotFound": "Verify that the file path is correct and the file exists.",
    "CreationError": "Check the operation parameters and try again.",
    "UpdateError": "Verify that the element can be updated and parameters are valid.",
    "SpecialIDNotAvailable": "Make sure you have performed the required operation before using this special ID.",
}


def create_markdown_response(
    session,
    message: str,
    element_id: Optional[str] = None,
    operation: Optional[str] = None,
    status: str = "success",
    show_context: bool = True,
    show_diff: bool = False,
    old_content: Optional[str] = None,
    new_content: Optional[str] = None,
    **extra_metadata
) -> str:
    """Create a Markdown-formatted response with ASCII visualization.

    Args:
        session: Session object (can be None for errors)
        message: Human-readable message
        element_id: Created/modified element ID
        operation: Operation name (e.g., "Insert Paragraph")
        status: "success" or "error"
        show_context: Whether to include document context visualization
        show_diff: Whether to show before/after diff
        old_content: Old content for diff (if show_diff=True)
        new_content: New content for diff (if show_diff=True)
        **extra_metadata: Additional metadata fields

    Returns:
        Markdown-formatted string
    """
    from docx_editor_skill.core.visualizer import DocumentVisualizer, DiffRenderer

    lines = []

    # Header
    operation_name = operation or "Operation"
    lines.append(f"# 操作结果: {operation_name}")
    lines.append("")

    # Status and metadata
    status_icon = "✅ Success" if status == "success" else "❌ Error"
    lines.append(f"**Status**: {status_icon}")

    if element_id:
        lines.append(f"**Element ID**: {element_id}")

    if operation:
        lines.append(f"**Operation**: {operation}")

    # Add extra metadata fields
    for key, value in extra_metadata.items():
        # Convert key from snake_case to Title Case with acronym handling
        words = key.split('_')
        title_words = []
        acronyms = {'id', 'url', 'html', 'xml', 'api'}

        for word in words:
            if word.lower() in acronyms:
                title_words.append(word.upper())
            else:
                title_words.append(word.capitalize())

        display_key = ' '.join(title_words)
        lines.append(f"**{display_key}**: {value}")

    # Show diff if requested
    if show_diff and old_content is not None and new_content is not None:
        lines.append("")
        lines.append("---")
        lines.append("")

        diff_renderer = DiffRenderer()
        element_type = extra_metadata.get('element_type', 'Paragraph')
        # Normalize non-string content for diff rendering
        if not isinstance(old_content, str):
            try:
                old_content = json.dumps(old_content, ensure_ascii=False, indent=2)
            except Exception:
                old_content = str(old_content)
        if not isinstance(new_content, str):
            try:
                new_content = json.dumps(new_content, ensure_ascii=False, indent=2)
            except Exception:
                new_content = str(new_content)
        diff_output = diff_renderer.render_diff(
            old_content, new_content,
            element_id or "unknown", element_type
        )
        lines.append(diff_output)

    # Show document context if requested and session is available
    if show_context and session and element_id:
        lines.append("")
        lines.append("---")
        lines.append("")
        lines.append("## 📄 Document Context")
        lines.append("")

        try:
            visualizer = DocumentVisualizer(session)
            context = visualizer.render_context(element_id)
            lines.append(context)
        except Exception as e:
            logger.warning(f"Failed to render context: {e}")
            lines.append(f"(Context rendering failed: {e})")

    return "\n".join(lines)


def create_error_response(message: str, error_type: Optional[str] = None) -> str:
    """Create an error response in Markdown format.

    Args:
        message: Error message
        error_type: Optional error type classification

    Returns:
        Markdown-formatted error string
    """
    lines = []
    lines.append("# 操作结果: Error")
    lines.append("")
    lines.append("**Status**: ❌ Error")

    if error_type:
        lines.append(f"**Error Type**: {error_type}")

    lines.append(f"**Message**: {message}")

    # Add suggestion if available for this error type
    if error_type and error_type in ERROR_SUGGESTIONS:
        lines.append("")
        lines.append(f"**Suggestion**: {ERROR_SUGGESTIONS[error_type]}")

    return "\n".join(lines)


# Legacy function names for backward compatibility during migration
# These will be removed after all tools are migrated
def create_success_response(
    message: str,
    element_id: Optional[str] = None,
    cursor: Optional[Any] = None,
    **extra_data
) -> str:
    """Legacy function - redirects to create_markdown_response.

    This function is kept for backward compatibility during migration.
    New code should use create_markdown_response directly.
    """
    logger.warning("create_success_response is deprecated, use create_markdown_response")
    # For now, return a simple markdown response without session context
    return create_markdown_response(
        session=None,
        message=message,
        element_id=element_id,
        show_context=False,
        **extra_data
    )


def create_context_aware_response(
    session,
    message: str,
    element_id: Optional[str] = None,
    include_cursor: bool = True,
    **extra_data
) -> str:
    """Legacy function - redirects to create_markdown_response.

    This function is kept for backward compatibility during migration.
    New code should use create_markdown_response directly.
    """
    logger.warning("create_context_aware_response is deprecated, use create_markdown_response")
    return create_markdown_response(
        session=session,
        message=message,
        element_id=element_id,
        show_context=include_cursor,
        **extra_data
    )


def create_change_tracked_response(
    session,
    message: str,
    element_id: Optional[str] = None,
    changes: Optional[Any] = None,
    commit_id: Optional[str] = None,
    include_cursor: bool = True,
    **extra_data
) -> str:
    """Legacy function - redirects to create_markdown_response.

    This function is kept for backward compatibility during migration.
    New code should use create_markdown_response directly.
    """
    logger.warning("create_change_tracked_response is deprecated, use create_markdown_response")

    # Extract old/new content from changes if available
    old_content = None
    new_content = None
    show_diff = False

    if changes and isinstance(changes, dict):
        old_content = changes.get('before')
        new_content = changes.get('after')
        show_diff = old_content is not None and new_content is not None

    if commit_id:
        extra_data['commit_id'] = commit_id

    return create_markdown_response(
        session=session,
        message=message,
        element_id=element_id,
        show_context=include_cursor,
        show_diff=show_diff,
        old_content=old_content,
        new_content=new_content,
        changes=changes,
        **extra_data
    )
