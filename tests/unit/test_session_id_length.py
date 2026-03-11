from tests.helpers.session_helpers import setup_active_session, teardown_active_session

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
    extract_error_message
)


def test_session_id_short_length():
    setup_active_session()
    try:
        from docx_editor_skill.core.global_state import global_state
        sid = global_state.active_session_id
        assert len(sid) <= 12
    finally:
        teardown_active_session()
