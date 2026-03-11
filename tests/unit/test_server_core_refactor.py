import pytest
from tests.helpers import (
    extract_session_id,
    extract_element_id,
    extract_metadata_field,
    is_success,
    is_error,
    create_session_with_file,
)
from tests.helpers.session_helpers import setup_active_session, teardown_active_session
import json
import re
from unittest.mock import MagicMock, patch
from docx_editor_skill.tools import (
    docx_insert_paragraph,
    docx_insert_run,
    docx_set_properties,
)
from docx_editor_skill.core.session import session_manager


def _extract_element_id(response):
    """Extract element_id from Markdown response."""
    # Try to extract from Markdown format: **Element ID**: para_xxx
    match = re.search(r'\*\*Element ID\*\*:\s*(\w+)', response)
    if match:
        return match.group(1)
    # Fallback: try JSON format (legacy)
    try:
        data = json.loads(response)
        if isinstance(data, dict) and "data" in data and "element_id" in data["data"]:
            return data["data"]["element_id"]
        return response
    except (json.JSONDecodeError, KeyError):
        return response


def test_create_with_autosave(tmp_path):
    test_file = str(tmp_path / "test.docx")
    with patch("docx_editor_skill.core.session.Document") as mock_doc:
        sid = create_session_with_file(test_file, auto_save=True)
        assert sid is not None
        session = session_manager.get_session(sid)
        assert session.auto_save is True
        assert session.file_path == test_file
        teardown_active_session()

def test_implicit_context_flow():
    # 1. Create session
    setup_active_session()
    try:
        from docx_editor_skill.core.global_state import global_state
        sid = global_state.active_session_id
        assert sid is not None
        session = session_manager.get_session(sid)

        # 2. Add paragraph (should set last_created_id)
        p_response = docx_insert_paragraph("Hello", position="end:document_body")
        p_id = _extract_element_id(p_response)
        clean_id = p_id.strip().split()[0]
        assert session.last_created_id == clean_id

        # 3. Add run using explicit position
        r_response = docx_insert_run(" World", position=f"inside:{p_id}")
        r_id = _extract_element_id(r_response)
        clean_rid = r_id.strip().split()[0]
        assert session.last_accessed_id == clean_rid

        # Verify run was added to paragraph
        para_obj = session.get_object(p_id)
        assert len(para_obj.runs) == 2 # "Hello" + " World" (first run created by add_paragraph)
    finally:
        teardown_active_session()

def test_set_properties_flow():
    setup_active_session()
    try:
        from docx_editor_skill.core.global_state import global_state
        sid = global_state.active_session_id
        p_response = docx_insert_paragraph("Test Prop", position="end:document_body")
        p_id = _extract_element_id(p_response)
        assert sid is not None

        # Set alignment via properties
        props = json.dumps({
            "paragraph_format": {"alignment": "center"}
        })
        docx_set_properties(props, element_id=p_id)

        from docx.enum.text import WD_ALIGN_PARAGRAPH
        para_obj = session_manager.get_session(sid).get_object(p_id)
        assert para_obj.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER
    finally:
        teardown_active_session()

def test_set_properties_implicit_context():
    setup_active_session()
    try:
        from docx_editor_skill.core.global_state import global_state
        sid = global_state.active_session_id
        assert sid is not None
        p_response = docx_insert_paragraph("Context Test", position="end:document_body")
        p_id = _extract_element_id(p_response)

        # Implicitly use p_id because it was last created/accessed
        props = json.dumps({"alignment": "right"}) # Shortcut property
        docx_set_properties(props) # No element_id

        from docx.enum.text import WD_ALIGN_PARAGRAPH
        para_obj = session_manager.get_session(sid).get_object(p_id)
        assert para_obj.alignment == WD_ALIGN_PARAGRAPH.RIGHT
    finally:
        teardown_active_session()
