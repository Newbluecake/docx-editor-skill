"""Unit tests for optimized content tools"""
import json
import pytest
from docx_editor_skill.tools.content_tools import (
    docx_read_content,
    docx_find_paragraphs,
    docx_extract_template_structure
)
from tests.helpers.session_helpers import setup_active_session, teardown_active_session
from docx_editor_skill.tools.paragraph_tools import docx_insert_paragraph, docx_insert_heading
from docx_editor_skill.tools.table_tools import docx_insert_table

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
    extract_find_paragraphs_results,
    extract_template_structure
)


def test_read_content_with_pagination():
    """Test reading content with pagination"""
    setup_active_session()
    try:
        # Add multiple paragraphs
        for i in range(20):
            docx_insert_paragraph(f"Paragraph {i}", position="end:document_body")

        # Read content as JSON to get reliable entry indices
        content_json = docx_read_content(max_paragraphs=50, return_json=True)
        data = json.loads(content_json)["data"]

        # Find index of our first inserted paragraph
        offset = 0
        found = False
        for i, entry in enumerate(data):
            if "text" in entry and "Paragraph 0" in entry["text"]:
                offset = i
                found = True
                break

        assert found, "Could not find 'Paragraph 0' in document content"

        # Verify first 5 of OUR paragraphs (offset to offset+5)
        # Note: data[offset] is Paragraph 0
        assert "Paragraph 0" in data[offset]["text"]
        assert "Paragraph 4" in data[offset+4]["text"]

        # Read paragraphs 5-10 (relative to our insertion)
        # We want to start at Paragraph 5, which is at index offset + 5
        start_index = offset + 5

        # Test text mode pagination
        content = docx_read_content(max_paragraphs=5, start_from=start_index)
        lines = [line for line in content.split("\n") if line.strip()]

        assert len(lines) == 5
        assert "Paragraph 5" in lines[0]
        assert "Paragraph 9" in lines[-1]

    finally:
        teardown_active_session()


def test_find_paragraphs_with_limit():
    """Test finding paragraphs with result limit"""
    setup_active_session()
    try:
        # Add many matching paragraphs
        for i in range(20):
            docx_insert_paragraph(f"Test paragraph {i}", position="end:document_body")

        # Find with limit
        matches_md = docx_find_paragraphs("Test", max_results=5)
        matches = extract_find_paragraphs_results(matches_md)

        assert len(matches) == 5

    finally:
        teardown_active_session()


def test_extract_template_structure_with_limits():
    """Test extracting structure with item limits"""
    setup_active_session()
    try:
        # Add content
        for i in range(5):
            docx_insert_heading(f"Heading {i}", position="end:document_body", level=1)
            docx_insert_paragraph(f"Content {i}", position="end:document_body")

        docx_insert_table(2, 2, position="end:document_body")

        # Extract with limits
        limits = json.dumps({"headings": 2, "paragraphs": 0, "tables": 1})
        structure_md = docx_extract_template_structure(
            max_items_per_type=limits,
            include_content=False
        )

        structure = extract_template_structure(structure_md)
        doc_structure = structure["document_structure"]

        # Count by type
        headings = [e for e in doc_structure if e["type"] == "heading"]
        paragraphs = [e for e in doc_structure if e["type"] == "paragraph"]
        tables = [e for e in doc_structure if e["type"] == "table"]

        assert len(headings) <= 2
        assert len(paragraphs) == 0
        assert len(tables) <= 1

    finally:
        teardown_active_session()


def test_extract_template_structure_no_content():
    """Test extracting structure without content"""
    setup_active_session()
    try:
        docx_insert_heading("Test Heading", position="end:document_body", level=1)
        docx_insert_paragraph("Long paragraph content" * 100, position="end:document_body")

        structure_md = docx_extract_template_structure(
            include_content=False
        )

        structure = extract_template_structure(structure_md)
        doc_structure = structure["document_structure"]

        # Check that text is truncated
        for item in doc_structure:
            if "text" in item:
                assert "[" in item["text"] and "chars]" in item["text"]

    finally:
        teardown_active_session()
