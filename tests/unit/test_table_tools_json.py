import json
import pytest
import sys
import os

# Add parent directory to path for helpers import
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from docx_editor_skill.tools.table_tools import (
    docx_insert_table,
    docx_get_table,
    docx_find_table,
    docx_get_cell,
    docx_insert_paragraph_to_cell,
    docx_insert_table_row,
    docx_insert_table_col,
    docx_fill_table,
    docx_copy_table
)
from tests.helpers.session_helpers import setup_active_session, teardown_active_session
from helpers import (
    extract_session_id,
    extract_element_id,
    extract_metadata_field,
    is_success,
    is_error,
    extract_error_message
)


def test_add_table_returns_json():
    """Test that docx_insert_table returns valid JSON."""
    setup_active_session()
    result = docx_insert_table(rows=2, cols=3, position="end:document_body")
    assert is_success(result)
    element_id = extract_element_id(result)
    assert element_id is not None
    assert element_id.startswith("table_")
    assert int(extract_metadata_field(result, "rows")) == 2
    assert int(extract_metadata_field(result, "cols")) == 3

    teardown_active_session()


def test_get_table_returns_json():
    """Test that docx_get_table returns valid JSON."""
    setup_active_session()
    docx_insert_table(2, 2, position="end:document_body")

    result = docx_get_table(0)

    assert is_success(result)
    element_id = extract_element_id(result)
    assert element_id is not None
    assert int(extract_metadata_field(result, "index")) == 0

    teardown_active_session()


def test_find_table_returns_json():
    """Test that docx_find_table returns valid JSON."""
    setup_active_session()
    # Create table with specific text
    table_id = extract_element_id(docx_insert_table(2, 2, position="end:document_body"))
    cell_id = extract_element_id(docx_get_cell(table_id, 0, 0))
    docx_insert_paragraph_to_cell("UniqueHeader", position=f"inside:{cell_id}")

    # Find it
    result = docx_find_table("UniqueHeader")

    assert is_success(result)
    assert extract_element_id(result) == table_id
    assert extract_metadata_field(result, "search_text") == "UniqueHeader"

    teardown_active_session()


def test_get_cell_returns_json():
    """Test that docx_get_cell returns valid JSON."""
    setup_active_session()
    table_id = extract_element_id(docx_insert_table(2, 2, position="end:document_body"))

    result = docx_get_cell(table_id, 0, 1)

    assert is_success(result)
    element_id = extract_element_id(result)
    assert element_id is not None
    assert element_id.startswith("cell_")
    assert int(extract_metadata_field(result, "row")) == 0
    assert int(extract_metadata_field(result, "col")) == 1

    teardown_active_session()


def test_add_paragraph_to_cell_returns_json():
    """Test that docx_insert_paragraph_to_cell returns valid JSON."""
    setup_active_session()
    table_id = extract_element_id(docx_insert_table(2, 2, position="end:document_body"))
    cell_id = extract_element_id(docx_get_cell(table_id, 0, 0))

    result = docx_insert_paragraph_to_cell("Cell Text", position=f"inside:{cell_id}")

    assert is_success(result)
    element_id = extract_element_id(result)
    assert element_id is not None

    teardown_active_session()


def test_add_table_row_returns_json():
    """Test that docx_insert_table_row returns valid JSON."""
    setup_active_session()
    table_id = extract_element_id(docx_insert_table(2, 2, position="end:document_body"))

    result = docx_insert_table_row(position=f"inside:{table_id}")

    assert is_success(result)
    assert int(extract_metadata_field(result, "new_row_count")) == 3

    teardown_active_session()


def test_add_table_col_returns_json():
    """Test that docx_insert_table_col returns valid JSON."""
    setup_active_session()
    table_id = extract_element_id(docx_insert_table(2, 2, position="end:document_body"))

    result = docx_insert_table_col(position=f"inside:{table_id}")

    assert is_success(result)
    assert int(extract_metadata_field(result, "new_col_count")) == 3

    teardown_active_session()


def test_fill_table_returns_json():
    """Test that docx_fill_table returns valid JSON."""
    setup_active_session()
    table_id = extract_element_id(docx_insert_table(2, 2, position="end:document_body"))

    data_json = json.dumps([["A1", "B1"], ["A2", "B2"]])
    result = docx_fill_table(data_json, table_id)

    assert is_success(result)
    assert int(extract_metadata_field(result, "rows_filled")) == 2

    teardown_active_session()


def test_fill_table_preserves_run_formatting_when_enabled():
    """docx_fill_table preserves existing run formatting when preserve_formatting=True."""
    setup_active_session()
    table_id = extract_element_id(docx_insert_table(1, 1, position="end:document_body"))

    # Seed formatting in the first cell (bold run)
    cell_id = extract_element_id(docx_get_cell(table_id, 0, 0))
    from docx_editor_skill.core.session import session_manager
    from docx_editor_skill.core.global_state import global_state
    session = session_manager.get_session(global_state.active_session_id)
    cell = session.get_object(cell_id)
    para = cell.paragraphs[0]
    run = para.add_run("Seed")
    run.font.bold = True

    # Fill table with preserve_formatting=True
    data_json = json.dumps([["New Value"]])
    result = docx_fill_table(data_json, table_id, preserve_formatting=True)

    assert is_success(result)
    assert extract_metadata_field(result, "preserve_formatting") is True

    # Verify formatting is retained on the new run
    updated_cell = session.get_object(table_id).rows[0].cells[0]
    updated_run = updated_cell.paragraphs[0].runs[0]
    assert updated_run.text == "New Value"
    assert updated_run.font.bold is True

    teardown_active_session()


def test_copy_table_returns_json():
    """Test that docx_copy_table returns valid JSON."""
    setup_active_session()
    table_id = extract_element_id(docx_insert_table(2, 2, position="end:document_body"))

    result = docx_copy_table(table_id, position="end:document_body")

    assert is_success(result)
    element_id = extract_element_id(result)
    assert element_id is not None
    assert element_id != table_id
    assert extract_metadata_field(result, "source_id") == table_id

    teardown_active_session()


def test_table_error_handling():
    """Test error handling in table tools."""
    setup_active_session()

    # Index out of bounds (no tables exist yet)
    result = docx_get_table(99)
    assert is_error(result)

    # Invalid JSON for fill_table
    table_id = extract_element_id(docx_insert_table(2, 2, position="end:document_body"))
    result = docx_fill_table("invalid_json", table_id)
    assert is_error(result)
    assert extract_metadata_field(result, "error_type") == "JSONDecodeError"

    teardown_active_session()


def test_table_no_active_session():
    """Test table tools with no active session."""
    teardown_active_session()

    result = docx_get_table(0)
    assert is_error(result)
    assert "NoActiveSession" in result
