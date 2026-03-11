import pytest
from unittest.mock import Mock
from docx.text.run import Run, Font
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_COLOR_INDEX
from docx_editor_skill.core.format_painter import FormatPainter

def test_copy_run_format_basic():
    painter = FormatPainter()
    src = Mock(spec=Run)
    tgt = Mock(spec=Run)

    # Setup source font
    src.font = Mock(spec=Font)
    src.font.name = "Arial"
    src.font.size = Pt(12)
    src.font.bold = True
    src.font.italic = False
    src.font.underline = True
    src.font.strike = False
    src.font.color.rgb = RGBColor(255, 0, 0)

    # Setup target font
    tgt.font = Mock(spec=Font)

    painter._copy_run_format(src, tgt)

    # Assertions
    assert tgt.font.name == "Arial"
    assert tgt.font.size == Pt(12)
    assert tgt.font.bold is True
    assert tgt.font.italic is False
    assert tgt.font.underline is True
    assert tgt.font.strike is False
    assert tgt.font.color.rgb == RGBColor(255, 0, 0)

def test_copy_run_format_partial():
    """Test copying when some source attributes are None (should copy None or keep target? Usually Format Painter overwrites with source state, even if None/Default)"""
    painter = FormatPainter()
    src = Mock(spec=Run)
    tgt = Mock(spec=Run)

    src.font = Mock(spec=Font)
    src.font.bold = None # Default

    tgt.font = Mock(spec=Font)
    tgt.font.bold = True

    painter._copy_run_format(src, tgt)

    assert tgt.font.bold is None

def test_copy_run_format_advanced_attributes():
    """Test superscript, subscript, highlight"""
    painter = FormatPainter()
    src = Mock(spec=Run)
    tgt = Mock(spec=Run)

    src.font = Mock(spec=Font)
    src.font.superscript = True
    src.font.subscript = False
    src.font.highlight_color = WD_COLOR_INDEX.YELLOW

    tgt.font = Mock(spec=Font)

    painter._copy_run_format(src, tgt)

    assert tgt.font.superscript is True
    assert tgt.font.subscript is False
    assert tgt.font.highlight_color == WD_COLOR_INDEX.YELLOW
