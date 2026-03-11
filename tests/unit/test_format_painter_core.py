import pytest
from unittest.mock import MagicMock, Mock
from docx_editor_skill.core.format_painter import FormatPainter
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docx.table import Table

def test_format_painter_dispatch_run():
    painter = FormatPainter()
    mock_src = Mock(spec=Run)
    mock_tgt = Mock(spec=Run)

    # Mock the internal methods to verify dispatch
    painter._copy_run_format = MagicMock()

    painter.copy_format(mock_src, mock_tgt)

    painter._copy_run_format.assert_called_once_with(mock_src, mock_tgt)

def test_format_painter_dispatch_paragraph():
    painter = FormatPainter()
    mock_src = Mock(spec=Paragraph)
    mock_tgt = Mock(spec=Paragraph)

    painter._copy_paragraph_format = MagicMock()

    painter.copy_format(mock_src, mock_tgt)

    painter._copy_paragraph_format.assert_called_once_with(mock_src, mock_tgt)

def test_format_painter_dispatch_table():
    painter = FormatPainter()
    mock_src = Mock(spec=Table)
    mock_tgt = Mock(spec=Table)

    painter._copy_table_format = MagicMock()

    painter.copy_format(mock_src, mock_tgt)

    painter._copy_table_format.assert_called_once_with(mock_src, mock_tgt)

def test_format_painter_unsupported_type():
    painter = FormatPainter()
    mock_src = "NotAnObject"  # type: ignore
    mock_tgt = "NotAnObject"  # type: ignore

    with pytest.raises(ValueError, match="Unsupported object type"):
        painter.copy_format(mock_src, mock_tgt)  # type: ignore
