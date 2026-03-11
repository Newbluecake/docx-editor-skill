import pytest
import json
from docx_editor_skill.core.session import session_manager
from docx_editor_skill.tools.paragraph_tools import docx_insert_paragraph
from docx_editor_skill.tools.table_tools import docx_insert_table
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

def test_add_table_position_after():
    setup_active_session()
    try:
        # Create Anchor Paragraph
        p1_resp = docx_insert_paragraph("Anchor", position="end:document_body")
        p1_id = extract_element_id(p1_resp)

        # Insert Table after Anchor
        t_resp = docx_insert_table(rows=2, cols=2, position=f"after:{p1_id}")
        t_id = extract_element_id(t_resp)

        # Verify Order
        session_id = global_state.active_session_id
        session = session_manager.get_session(session_id)
        # Document elements are stored in body._element
        # We can iterate block items to check order
        # Or just check if table is after paragraph
        doc_body = session.document._body
        # python-docx doesn't expose a unified list of paragraphs and tables easily in .elements
        # But we can check xml path or child order
        children = list(doc_body._element.iterchildren())
        assert len(children) >= 2
        # Expect: P (Anchor), Tbl
        assert children[0].tag.endswith('p')
        assert children[1].tag.endswith('tbl')

        # Verify success
        assert is_success(t_resp)
    finally:
        teardown_active_session()

def test_add_table_position_start():
    setup_active_session()
    try:
        docx_insert_paragraph("Existing", position="end:document_body")

        # Insert Table at start
        docx_insert_table(1, 1, position="start:document_body")

        session_id = global_state.active_session_id
        session = session_manager.get_session(session_id)
        children = list(session.document._body._element.iterchildren())
        # Expect: Tbl, P
        assert children[0].tag.endswith('tbl')
        assert children[1].tag.endswith('p')
    finally:
        teardown_active_session()
