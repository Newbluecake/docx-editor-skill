import json
import os
import tempfile

from tests.helpers.session_helpers import setup_active_session, teardown_active_session
from docx_editor_skill.tools.paragraph_tools import docx_insert_paragraph
from docx_editor_skill.tools.content_tools import (
    docx_read_content,
    docx_find_paragraphs,
)

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
    extract_find_paragraphs_results
)


def test_read_content_return_json_and_ids():
    setup_active_session()
    try:
        docx_insert_paragraph("alpha", position="end:document_body")
        docx_insert_paragraph("beta", position="end:document_body")

        result = json.loads(docx_read_content(return_json=True, include_ids=True))

        assert result["status"] == "success"
        assert result["count"] == 2
        assert all("id" in entry and entry["text"] for entry in result["data"])
    finally:
        teardown_active_session()


def test_find_paragraphs_with_context():
    setup_active_session()
    try:
        docx_insert_paragraph("one", position="end:document_body")
        docx_insert_paragraph("two match", position="end:document_body")
        docx_insert_paragraph("three", position="end:document_body")

        matches_md = docx_find_paragraphs("match", max_results=2, return_context=True, context_span=1)
        matches = extract_find_paragraphs_results(matches_md)
        assert len(matches) == 1
        entry = matches[0]
        assert entry["text"] == "two match"
        assert entry["context_before"] == ["one"]
        assert entry["context_after"] == ["three"]
    finally:
        teardown_active_session()
