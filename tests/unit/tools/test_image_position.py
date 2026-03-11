import pytest
import tempfile
import os
import json
from docx_editor_skill.core.session import session_manager
from docx_editor_skill.tools.advanced_tools import docx_insert_image
from docx_editor_skill.tools.paragraph_tools import docx_insert_paragraph
from docx_editor_skill.core.global_state import global_state

# Add parent directory to path for helpers import
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../..')))

from helpers import (
    extract_session_id,
    extract_element_id,
    extract_metadata_field,
    is_success,
    is_error,
    extract_error_message
)
from tests.helpers.session_helpers import setup_active_session, teardown_active_session

# Minimal 1x1 PNG
PNG_DATA = b'\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x06\x00\x00\x00\x1f\x15\xc4\x89\x00\x00\x00\rIDATx\x9cc\xf8\xff\xff\x3f\x03\x00\x08\xfc\x02\xfe\xa7\x9a\xa0\xa0\x00\x00\x00\x00IEND\xaeB`\x82'

@pytest.fixture
def temp_image():
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
        f.write(PNG_DATA)
        path = f.name
    yield path
    if os.path.exists(path):
        os.unlink(path)

def test_insert_image_position_after(temp_image):
    setup_active_session()
    try:
        # Create Anchor
        p1_resp = docx_insert_paragraph("Anchor", position="end:document_body")
        p1_id = extract_element_id(p1_resp)

        # Insert Image after Anchor
        # This creates a new paragraph containing the image
        img_resp = docx_insert_image(temp_image, position=f"after:{p1_id}")
        img_id = extract_element_id(img_resp)

        session_id = global_state.active_session_id
        session = session_manager.get_session(session_id)
        # Filter for paragraphs only (ignore sectPr)
        children = [c for c in session.document._body._element.iterchildren() if c.tag.endswith('p')]

        # Expect: Anchor (P), Image (P)
        assert len(children) == 2
        assert children[0].text == "Anchor"

        # Verify success
        assert is_success(img_resp)
    finally:
        teardown_active_session()

def test_insert_image_position_start(temp_image):
    setup_active_session()
    try:
        docx_insert_paragraph("Existing", position="end:document_body")

        # Insert at start
        img_resp = docx_insert_image(temp_image, position="start:document_body")

        session_id = global_state.active_session_id
        session = session_manager.get_session(session_id)
        # Filter for paragraphs only (ignore sectPr)
        children = [c for c in session.document._body._element.iterchildren() if c.tag.endswith('p')]

        # Expect: Image (P), Existing (P)
        assert len(children) == 2
        assert children[1].text == "Existing"
    finally:
        teardown_active_session()

def test_insert_image_in_paragraph_error_handling(temp_image):
    # Test trying to insert into an incompatible parent or with invalid position
    setup_active_session()
    try:
        # Invalid position format
        resp = docx_insert_image(temp_image, position="invalid:format")
        assert is_error(resp)
        assert extract_metadata_field(resp, "error_type") == "ValidationError"
    finally:
        teardown_active_session()
