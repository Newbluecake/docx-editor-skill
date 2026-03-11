"""Pytest configuration to ensure test helpers are importable.

This adds the repository root and the tests directory to sys.path so that
`tests.helpers` can be imported consistently across all test modules.
"""

import os
import sys
import json
import pytest
from docx_editor_skill.core.global_state import global_state

from tests.helpers import (
    extract_status,
    extract_all_metadata,
    extract_error_message,
)


@pytest.fixture(autouse=True)
def reset_global_state():
    """Reset global state before each test to ensure isolation."""
    global_state.clear()
    yield
    global_state.clear()


@pytest.fixture
def active_session():
    """Create and teardown an active session for testing.

    This fixture provides a session_id that is automatically set as the
    active session in global_state. Tests can use this fixture instead
    of manually calling setup_active_session/teardown_active_session.

    Usage:
        def test_something(active_session):
            # active_session is already set in global_state
            result = docx_insert_paragraph("text", position="end:document_body")
            # ... assertions ...
            # teardown happens automatically
    """
    from tests.helpers.session_helpers import setup_active_session, teardown_active_session

    session_id = setup_active_session()
    yield session_id
    teardown_active_session()


ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
TESTS_DIR = os.path.abspath(os.path.dirname(__file__))

for path in (ROOT_DIR, TESTS_DIR):
    if path not in sys.path:
        sys.path.insert(0, path)


# Backward compatibility: allow existing tests that still call json.loads on
# tool responses to work with the new Markdown response format by converting
# Markdown into a lightweight dict structure.
_original_json_loads = json.loads


def _markdown_aware_loads(s, *args, **kwargs):
    if isinstance(s, str):
        try:
            return _original_json_loads(s, *args, **kwargs)
        except json.JSONDecodeError as e:
            # If this looks like a markdown response, parse metadata instead of raising
            status = extract_status(s)
            if status != "unknown" or "**Status**" in s or "操作结果" in s:
                metadata = extract_all_metadata(s)
                message = "" if status == "success" else (extract_error_message(s) or "")
                return {
                    "status": status,
                    "message": message,
                    "data": metadata,
                }
            # Not markdown-like; propagate original JSON error
            raise e
    return _original_json_loads(s, *args, **kwargs)


json.loads = _markdown_aware_loads
