"""Unit tests for composite tools"""
import json
import pytest
from docx_editor_skill.tools.composite_tools import (
    docx_insert_formatted_paragraph,
    docx_quick_edit,
    docx_get_structure_summary,
    docx_smart_fill_table,
    docx_format_range
)
from tests.helpers.session_helpers import setup_active_session, teardown_active_session
from docx_editor_skill.tools.paragraph_tools import docx_insert_paragraph
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
    extract_json_from_markdown
)


def test_add_formatted_paragraph():
    """Test creating formatted paragraph in one step"""
    setup_active_session()
    try:
        # Create formatted paragraph
        result = docx_insert_formatted_paragraph(
            "Test Text",
            position="end:document_body",
            bold=True,
            size=14,
            color_hex="FF0000",
            alignment="center"
        )

        para_id = extract_element_id(result)
        assert para_id.startswith("para_")

    finally:
        teardown_active_session()


def test_quick_edit():
    """Test quick edit functionality"""
    setup_active_session()
    try:
        # Add some paragraphs
        docx_insert_paragraph("This is a test paragraph", position="end:document_body")
        docx_insert_paragraph("Another test paragraph", position="end:document_body")

        # Quick edit
        result_md = docx_quick_edit(
            "test",
            new_text="modified",
            bold=True
        )

        # Extract data from Markdown response
        result = extract_json_from_markdown(result_md)
        assert result["modified_count"] == 2
        assert len(result["paragraph_ids"]) == 2

    finally:
        teardown_active_session()


def test_get_structure_summary():
    """Test lightweight structure extraction"""
    setup_active_session()
    try:
        # Add some content
        docx_insert_paragraph("Test", position="end:document_body", style="Heading 1")
        docx_insert_paragraph("Body text", position="end:document_body")
        docx_insert_table(2, 2, position="end:document_body")

        # Get summary
        summary_md = docx_get_structure_summary(
            max_headings=10,
            max_tables=5,
            max_paragraphs=0
        )

        # Extract data from Markdown response
        summary = extract_json_from_markdown(summary_md)
        assert "headings" in summary
        assert "tables" in summary
        assert summary["summary"]["total_headings"] >= 1
        assert summary["summary"]["total_tables"] >= 1

    finally:
        teardown_active_session()


def test_smart_fill_table():
    """Test smart table filling"""
    setup_active_session()
    try:
        # Create table
        table_response = docx_insert_table(2, 3, position="end:document_body")
        table_id = extract_element_id(table_response)

        # Fill with data
        data = json.dumps([
            ["Name", "Age", "City"],
            ["Alice", "30", "NYC"],
            ["Bob", "25", "LA"]
        ])

        result = docx_smart_fill_table(
            table_id,  # Use table ID
            data,
            has_header=True,
            auto_resize=True
        )

        assert is_success(result)
        assert extract_metadata_field(result, "rows_filled") == 3

    finally:
        teardown_active_session()


def test_format_range():
    """Test formatting a range of paragraphs"""
    setup_active_session()
    try:
        # Add paragraphs
        docx_insert_paragraph("Start marker", position="end:document_body")
        docx_insert_paragraph("Middle content", position="end:document_body")
        docx_insert_paragraph("End marker", position="end:document_body")

        # Format range
        result_md = docx_format_range(
            "Start marker",
            "End marker",
            bold=True,
            size=14
        )

        # Extract data from Markdown response
        result = extract_json_from_markdown(result_md)
        assert result["formatted_count"] == 3

    finally:
        teardown_active_session()


def test_quick_edit_no_matches():
    """Test quick edit with no matches"""
    setup_active_session()
    try:
        docx_insert_paragraph("Test paragraph", position="end:document_body")

        result_md = docx_quick_edit(
            "nonexistent",
            new_text="modified"
        )

        # Extract data from Markdown response
        result = extract_json_from_markdown(result_md)
        assert result["modified_count"] == 0

    finally:
        teardown_active_session()


def test_smart_fill_table_with_table_id():
    """Test smart table filling using table ID instead of index"""
    setup_active_session()
    try:
        # Create table and get its ID
        table_response = docx_insert_table(2, 3, position="end:document_body")
        table_id = extract_element_id(table_response)

        # Fill with data using table ID
        data = json.dumps([
            ["Name", "Age", "City"],
            ["Alice", "30", "NYC"],
            ["Bob", "25", "LA"],
            ["Charlie", "35", "SF"]
        ])

        result = docx_smart_fill_table(
            table_id,  # Use table ID directly
            data,
            has_header=True,
            auto_resize=True
        )

        assert is_success(result)
        assert extract_metadata_field(result, "rows_filled") == 4
        assert extract_metadata_field(result, "rows_added") == 2  # Should add 2 rows (started with 2, need 4 total)

    finally:
        teardown_active_session()
