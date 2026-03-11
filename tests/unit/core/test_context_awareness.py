"""Unit tests for cursor context awareness functionality."""

import pytest
from docx_editor_skill.core.session import SessionManager


@pytest.fixture
def session_manager():
    return SessionManager()


def test_context_empty_doc(session_manager):
    """Test context in empty document."""
    session_id = session_manager.create_session()
    session = session_manager.get_session(session_id)

    context = session.get_cursor_context()
    assert "empty document" in context.lower()


def test_context_start_end(session_manager):
    """Test context at document start and end."""
    session_id = session_manager.create_session()
    session = session_manager.get_session(session_id)

    # Add paragraphs
    para1 = session.document.add_paragraph("First")
    para1_id = session.register_object(para1, "para")
    session.cursor.element_id = para1_id

    para2 = session.document.add_paragraph("Second")
    para2_id = session.register_object(para2, "para")

    para3 = session.document.add_paragraph("Third")
    para3_id = session.register_object(para3, "para")
    session.cursor.element_id = para3_id

    # Test at end
    context = session.get_cursor_context()
    assert "Document End" in context or para3_id in context


def test_context_middle(session_manager):
    """Test context in middle of document."""
    session_id = session_manager.create_session()
    session = session_manager.get_session(session_id)

    # Add multiple paragraphs
    para1 = session.document.add_paragraph("Para 1")
    session.register_object(para1, "para")

    para2 = session.document.add_paragraph("Para 2")
    para2_id = session.register_object(para2, "para")

    para3 = session.document.add_paragraph("Para 3")
    session.register_object(para3, "para")

    # Set cursor to middle
    session.cursor.element_id = para2_id

    context = session.get_cursor_context(num_before=1, num_after=1)
    assert "[-1]" in context
    assert "[+1]" in context
    assert "[Current]" in context


def test_context_nested_cell(session_manager):
    """Test context inside table cell."""
    session_id = session_manager.create_session()
    session = session_manager.get_session(session_id)

    # Add table
    table = session.document.add_table(rows=2, cols=2)
    table_id = session.register_object(table, "table")

    # Add paragraph to cell
    cell = table.rows[0].cells[0]
    cell_id = session.register_object(cell, "cell")
    para = cell.add_paragraph("Cell content")
    para_id = session.register_object(para, "para")

    # Set cursor in cell
    session.cursor.element_id = para_id
    session.cursor.parent_id = cell_id

    context = session.get_cursor_context()
    assert "Parent:" in context
    assert cell_id in context


def test_context_truncation(session_manager):
    """Test long text truncation."""
    session_id = session_manager.create_session()
    session = session_manager.get_session(session_id)

    # Add paragraph with long text
    long_text = "A" * 100
    para = session.document.add_paragraph(long_text)
    para_id = session.register_object(para, "para")
    session.cursor.element_id = para_id

    context = session.get_cursor_context()
    assert "..." in context
    assert len([line for line in context.split('\n') if 'A' * 100 in line]) == 0
