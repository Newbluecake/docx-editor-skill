import pytest
import os
from docx_editor_skill.tools import docx_save, docx_read_content, docx_insert_paragraph, docx_find_paragraphs
from docx import Document
import json

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
    create_session_with_file,
)
from tests.helpers.session_helpers import setup_active_session, teardown_active_session

class TestLoadExisting:
    def test_create_with_existing_file(self, tmp_path):
        # Setup: Create a real docx file
        file_path = tmp_path / "test_existing.docx"
        doc = Document()
        doc.add_paragraph("Existing Content")
        doc.save(str(file_path))

        # Test: Load it via create_session_with_file
        session_id = create_session_with_file(str(file_path))
        # create_session_with_file returns session_id string directly
        assert session_id is not None
        assert len(session_id) > 0

        # Verify content can be read (using our new tool later, or just save and check)
        # For now, let's verify we can add to it and save
        docx_insert_paragraph("New Content", position="end:document_body")

        output_path = tmp_path / "test_existing_modified.docx"
        docx_save(str(output_path))

        # Verify output
        doc2 = Document(str(output_path))
        texts = [p.text for p in doc2.paragraphs]
        assert "Existing Content" in texts
        assert "New Content" in texts

        teardown_active_session()

    def test_create_with_nonexistent_file(self, tmp_path):
        # Test that creating a session with a non-existent file creates a new document
        # intended to be saved to that path later.
        target_path = tmp_path / "new_doc.docx"
        session_id = create_session_with_file(str(target_path))
        # create_session_with_file returns session_id string directly
        assert session_id is not None
        assert len(session_id) > 0

        # Verify it's an empty document (or default styles)
        content = docx_read_content()
        assert content == "[Empty Document]"

        teardown_active_session()

    def test_read_content(self, tmp_path):
        # Setup
        file_path = tmp_path / "test_read.docx"
        doc = Document()
        doc.add_paragraph("Line 1")
        doc.add_paragraph("Line 2")
        doc.save(str(file_path))

        session_id = create_session_with_file(str(file_path))
        # Test read
        content = docx_read_content()
        assert "Line 1" in content
        assert "Line 2" in content

        teardown_active_session()

    def test_find_paragraphs(self, tmp_path):
        # Setup
        file_path = tmp_path / "test_find.docx"
        doc = Document()
        doc.add_paragraph("First paragraph")
        doc.add_paragraph("Target paragraph here")
        doc.add_paragraph("Last paragraph")
        doc.save(str(file_path))

        session_id = create_session_with_file(str(file_path))
        # Test find - docx_find_paragraphs now returns Markdown
        result_md = docx_find_paragraphs("Target")

        # Parse Markdown response
        assert "# Found 1 matching paragraph(s)" in result_md
        assert "Target paragraph here" in result_md
        assert "para_" in result_md

        # Test find - no match
        result_md = docx_find_paragraphs("Nonexistent")
        assert "No matching paragraphs found" in result_md

        # Test case insensitivity
        result_md = docx_find_paragraphs("target")
        assert "# Found 1 matching paragraph(s)" in result_md

        teardown_active_session()
