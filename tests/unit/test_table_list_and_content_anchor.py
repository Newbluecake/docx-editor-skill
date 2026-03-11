import json

from tests.helpers.session_helpers import setup_active_session, teardown_active_session
from docx_editor_skill.tools.paragraph_tools import docx_insert_paragraph
from docx_editor_skill.tools.table_tools import docx_insert_table, docx_list_tables
from docx_editor_skill.tools.content_tools import docx_read_content

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


def test_list_tables_with_start_element_and_max_results():
    setup_active_session()
    try:
        # Add paragraphs and tables
        p1 = docx_insert_paragraph("before tables", position="end:document_body")
        t1 = docx_insert_table(rows=1, cols=1, position="end:document_body")
        p2 = docx_insert_paragraph("middle", position="end:document_body")
        t2 = docx_insert_table(rows=2, cols=2, position="end:document_body")
        t3 = docx_insert_table(rows=3, cols=1, position="end:document_body")

        # Start listing after t1
        t1_id = extract_element_id(t1) if isinstance(t1, str) else t1
        resp = json.loads(docx_list_tables(max_results=1, start_element_id=t1_id))
        assert resp["status"] == "success"
        assert resp["data"]["count"] == 1
        first = resp["data"]["tables"][0]
        assert first["rows"] == 2 and first["cols"] == 2

    finally:
        teardown_active_session()


def test_read_content_start_element_id_skips_prior_blocks():
    setup_active_session()
    try:
        docx_insert_paragraph("p0", position="end:document_body")
        anchor = docx_insert_paragraph("anchor", position="end:document_body")
        docx_insert_paragraph("p1", position="end:document_body")
        docx_insert_paragraph("p2", position="end:document_body")

        anchor_id = extract_element_id(anchor) if isinstance(anchor, str) else anchor
        result_json = docx_read_content(start_element_id=anchor_id, return_json=True)
        result = json.loads(result_json)
        texts = [e["text"] for e in result["data"]]
        assert texts[0] == "p1"
        assert len(texts) == 2
    finally:
        teardown_active_session()
