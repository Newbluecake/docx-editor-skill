import pytest
from unittest.mock import Mock, MagicMock
from docx.text.paragraph import Paragraph
from docx.text.run import Run, Font
from docx.styles.style import _ParagraphStyle
from docx_editor_skill.core.format_painter import FormatPainter

def test_smart_match_para_to_run():
    """
    Test copying from Paragraph to Run.
    Should apply the Paragraph's style font properties to the Run.
    """
    painter = FormatPainter()
    src = Mock(spec=Paragraph)
    tgt = Mock(spec=Run)

    # Mock Source Style
    style = Mock(spec=_ParagraphStyle)
    style.font = Mock(spec=Font)
    style.font.bold = True
    style.font.name = "Arial"
    src.style = style

    tgt.font = Mock(spec=Font)

    painter.copy_format(src, tgt)

    # Target Run should inherit style font properties
    assert tgt.font.bold is True
    assert tgt.font.name == "Arial"

def test_smart_match_run_to_para():
    """
    Test copying from Run to Paragraph.
    Should apply Run's font properties to all runs in the Paragraph.
    """
    painter = FormatPainter()
    src = Mock(spec=Run)
    tgt = Mock(spec=Paragraph)

    # Mock Source Run
    src.font = Mock(spec=Font)
    src.font.italic = True
    src.font.size = 240000 # EMU or something, strictly it's an object

    # Mock Target Paragraph with multiple runs
    run1 = Mock(spec=Run)
    run1.font = Mock(spec=Font)
    run2 = Mock(spec=Run)
    run2.font = Mock(spec=Font)
    tgt.runs = [run1, run2]

    painter.copy_format(src, tgt)

    # Both runs in paragraph should get the formatting
    assert run1.font.italic is True
    assert run1.font.size == 240000
    assert run2.font.italic is True
    assert run2.font.size == 240000
