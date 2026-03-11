"""Integration test for docx_extract_template_structure tool."""
import json
from docx import Document
from docx_editor_skill.tools import docx_extract_template_structure

# Add parent directory to path for helpers import
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from helpers import (
    extract_session_id,
    extract_element_id,
    extract_metadata_field,
    is_success,
    is_error
)
from tests.helpers.session_helpers import setup_active_session, teardown_active_session


def test_extract_template_structure_integration():
    """Test the MCP tool integration."""
    # Create a session with a simple template
    setup_active_session()
    try:
        # Create the document using the docx API directly
        from docx_editor_skill.core.session import session_manager
        from docx_editor_skill.core.global_state import global_state
        from tests.helpers import extract_template_structure
        session = session_manager.get_session(global_state.active_session_id)
        doc = session.document

        # Add elements
        doc.add_heading("Test Heading", level=1)
        doc.add_paragraph("Test paragraph")

        table = doc.add_table(rows=2, cols=2)
        for i, cell in enumerate(table.rows[0].cells):
            cell.text = f"Header {i + 1}"
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.bold = True

        # Extract structure
        result_markdown = docx_extract_template_structure()
        result = extract_template_structure(result_markdown)

        # Verify structure
        assert "metadata" in result
        assert "document_structure" in result
        assert len(result["document_structure"]) == 3

        # Verify element types
        assert result["document_structure"][0]["type"] == "heading"
        assert result["document_structure"][1]["type"] == "paragraph"
        assert result["document_structure"][2]["type"] == "table"

        # Verify table details
        table_data = result["document_structure"][2]
        assert table_data["rows"] == 2
        assert table_data["cols"] == 2
        assert table_data["headers"] == ["Header 1", "Header 2"]
    finally:
        # Clean up
        teardown_active_session()


def test_extract_template_structure_error_handling():
    """Test error handling for no active session."""
    teardown_active_session()
    result = docx_extract_template_structure()
    assert is_error(result)
    assert "NoActiveSession" in result or "no active session" in result.lower()
