"""Global state management for active file tracking.

This module provides a thread-safe global state manager that tracks the currently
active file and session. It is designed to be used in HTTP mode where concurrent
requests may access the state.

Thread Safety:
    - All property access is protected by threading.Lock
    - Use the atomic() context manager for compound operations
    - Compatible with both STDIO and HTTP transports

Example:
    >>> from docx_editor_skill.core.global_state import global_state
    >>>
    >>> # Simple access
    >>> global_state.active_file = "/path/to/doc.docx"
    >>> print(global_state.active_file)
    >>>
    >>> # Compound operation
    >>> with global_state.atomic():
    >>>     if global_state.active_session_id:
    >>>         # Close session
    >>>         global_state.active_session_id = None
"""

import threading
import logging
from typing import Optional
from contextlib import contextmanager

logger = logging.getLogger(__name__)


class GlobalState:
    """Thread-safe global state manager for active file and session tracking.

    This class maintains the global state of the currently active file and session.
    All property access is protected by a threading lock to ensure thread safety
    in HTTP mode.

    Attributes:
        active_file: Path to the currently active .docx file (or None)
        active_session_id: ID of the currently active session (or None)
    """

    def __init__(self):
        """Initialize the global state with thread safety."""
        self._lock = threading.RLock()
        self._active_file: Optional[str] = None
        self._active_session_id: Optional[str] = None
        logger.debug("GlobalState initialized")

    @contextmanager
    def atomic(self):
        """Context manager for atomic compound operations.

        Use this when you need to perform multiple state checks/updates atomically.

        Example:
            >>> with global_state.atomic():
            >>>     if global_state.active_file:
            >>>         # Update session
            >>>         global_state.active_session_id = "new-id"

        Yields:
            None: The lock is held for the duration of the context
        """
        with self._lock:
            yield

    @property
    def active_file(self) -> Optional[str]:
        """Get the currently active file path.

        Returns:
            Optional[str]: Path to active file, or None if no file is active
        """
        with self._lock:
            return self._active_file

    @active_file.setter
    def active_file(self, value: Optional[str]):
        """Set the currently active file path.

        Args:
            value: Path to the file to make active, or None to clear
        """
        with self._lock:
            old_value = self._active_file
            self._active_file = value
            if old_value != value:
                logger.info(f"Active file changed: {old_value} -> {value}")

    @property
    def active_session_id(self) -> Optional[str]:
        """Get the currently active session ID.

        Returns:
            Optional[str]: Active session ID, or None if no session is active
        """
        with self._lock:
            return self._active_session_id

    @active_session_id.setter
    def active_session_id(self, value: Optional[str]):
        """Set the currently active session ID.

        Args:
            value: Session ID to make active, or None to clear
        """
        with self._lock:
            old_value = self._active_session_id
            self._active_session_id = value
            if old_value != value:
                logger.info(f"Active session changed: {old_value} -> {value}")

    def clear(self):
        """Clear all state atomically.

        This is useful when resetting the server state or during testing.
        """
        with self._lock:
            self._active_file = None
            self._active_session_id = None
            logger.info("Global state cleared")

    def get_status(self) -> dict:
        """Get current status snapshot atomically.

        Returns:
            dict: Status dictionary with currentFile and sessionId keys
        """
        with self._lock:
            return {
                "currentFile": self._active_file,
                "sessionId": self._active_session_id
            }

    def set_active_file(self, file_path: str):
        """Set the active file path (legacy API for compatibility).

        Args:
            file_path: Path to the file to make active
        """
        self.active_file = file_path

    def clear_active_file(self):
        """Clear the active file (legacy API for compatibility)."""
        self.clear()


# Global singleton instance
global_state = GlobalState()
