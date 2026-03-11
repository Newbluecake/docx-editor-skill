"""Utility decorators for MCP tools"""
import logging
from functools import wraps
from typing import Callable

logger = logging.getLogger(__name__)


def require_session(func: Callable) -> Callable:
    """Decorator: Validate session_id and inject session object

    Transforms the first parameter from session_id (str) to session (Session object).
    Raises ValueError if session not found.
    """
    @wraps(func)
    def wrapper(session_id: str, *args, **kwargs):
        from docx_editor_skill.core.session import session_manager

        session = session_manager.get_session(session_id)
        if not session:
            logger.error(f"{func.__name__} failed: Session {session_id} not found")
            raise ValueError(f"Session {session_id} not found or expired")

        return func(session, *args, **kwargs)

    return wrapper


def log_tool_call(func: Callable) -> Callable:
    """Decorator: Log tool invocations and errors"""
    @wraps(func)
    def wrapper(*args, **kwargs):
        logger.debug(f"{func.__name__} called with args={args}, kwargs={kwargs}")
        try:
            result = func(*args, **kwargs)
            logger.debug(f"{func.__name__} success")
            return result
        except Exception as e:
            logger.exception(f"{func.__name__} failed: {e}")
            raise

    return wrapper
