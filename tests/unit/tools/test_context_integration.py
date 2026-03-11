"""Test context awareness integration with tools."""

import pytest
import json
import sys
import os

# Add parent directory to path for helpers import
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
from docx_editor_skill.tools.paragraph_tools import docx_insert_paragraph, docx_insert_heading
from docx_editor_skill.tools.run_tools import docx_insert_run
from docx_editor_skill.tools.table_tools import docx_insert_table, docx_insert_paragraph_to_cell, docx_get_cell


def test_paragraph_context():
    """Test that add_paragraph returns context."""
    setup_active_session()
    try:
        # Add first paragraph
        result1 = docx_insert_paragraph("First paragraph", position="end:document_body")
        assert is_success(result1)
        element_id1 = extract_element_id(result1)
        assert "para_" in element_id1
        # Document context is included in body; cursor metadata may be visual only
        assert "Document Context" in result1

        # Add second paragraph
        result2 = docx_insert_paragraph("Second paragraph", position="end:document_body")
        assert is_success(result2)
        assert "Document Context" in result2
    finally:
        teardown_active_session()


def test_heading_context():
    """Test that add_heading returns context."""
    setup_active_session()
    try:
        result = docx_insert_heading("My Heading", position="end:document_body", level=1)
        assert is_success(result)
        element_id = extract_element_id(result)
        assert "para_" in element_id
    finally:
        teardown_active_session()


def test_run_context():
    """Test that add_run returns context."""
    setup_active_session()
    try:
        para_result = docx_insert_paragraph("", position="end:document_body")
        para_id = extract_element_id(para_result)

        result = docx_insert_run("Run text", position=f"inside:{para_id}")
        assert is_success(result)
        element_id = extract_element_id(result)
        assert "run_" in element_id
    finally:
        teardown_active_session()


def test_table_context():
    """Test that add_table returns context."""
    setup_active_session()
    try:
        # Add paragraph first
        docx_insert_paragraph("Before table", position="end:document_body")

        # Add table
        result = docx_insert_table(rows=2, cols=2, position="end:document_body")
        assert is_success(result)
        element_id = extract_element_id(result)
        assert "table_" in element_id
    finally:
        teardown_active_session()


def test_table_cell_context():
    """Test context in table cells."""
    setup_active_session()
    try:
        table_result = docx_insert_table(rows=2, cols=2, position="end:document_body")
        table_id = extract_element_id(table_result)

        cell_result = docx_get_cell(table_id, 0, 0)
        cell_id = extract_element_id(cell_result)

        result = docx_insert_paragraph_to_cell("Cell content", position=f"inside:{cell_id}")
        assert is_success(result)
        element_id = extract_element_id(result)
        assert "para_" in element_id
    finally:
        teardown_active_session()


def test_advanced_tools_context():
    """Test context in advanced tools."""
    from docx_editor_skill.tools.advanced_tools import docx_insert_image
    import tempfile
    import os

    setup_active_session()
    try:
        # Create a minimal valid PNG file
        minimal_png = b'\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x06\x00\x00\x00\x1f\x15\xc4\x89\x00\x00\x00\nIDATx\x9cc\x00\x01\x00\x00\x05\x00\x01\r\n-\xb4\x00\x00\x00\x00IEND\xaeB`\x82'

        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
            f.write(minimal_png)
            img_path = f.name

        try:
            result = docx_insert_image(img_path, position="end:document_body")
            assert is_success(result)
            element_id = extract_element_id(result)
            assert "para_" in element_id or "run_" in element_id
        finally:
            if os.path.exists(img_path):
                os.remove(img_path)
    finally:
        teardown_active_session()
