"""Tests for advanced font attributes in FormatPainter."""
import pytest
from unittest.mock import Mock
from docx.text.run import Run, Font
from docx_editor_skill.core.format_painter import FormatPainter


def test_copy_advanced_font_attributes():
    """Test copying advanced font properties: all_caps, small_caps, double_strike, etc."""
    painter = FormatPainter()
    src = Mock(spec=Run)
    tgt = Mock(spec=Run)

    # Setup source font with advanced attributes
    src.font = Mock(spec=Font)
    src.font.all_caps = True
    src.font.small_caps = False
    src.font.double_strike = True
    src.font.shadow = True
    src.font.outline = False
    src.font.rtl = True
    src.font.emboss = False
    src.font.imprint = True

    # Setup target font
    tgt.font = Mock(spec=Font)

    painter._copy_run_format(src, tgt)

    # Verify all advanced attributes were copied
    assert tgt.font.all_caps is True
    assert tgt.font.small_caps is False
    assert tgt.font.double_strike is True
    assert tgt.font.shadow is True
    assert tgt.font.outline is False
    assert tgt.font.rtl is True
    assert tgt.font.emboss is False
    assert tgt.font.imprint is True


def test_copy_font_when_source_attributes_not_available():
    """Test that copying doesn't fail when source font doesn't have advanced attributes."""
    painter = FormatPainter()
    src = Mock(spec=Run)
    tgt = Mock(spec=Run)

    # Source font without advanced attributes (old python-docx version)
    src.font = Mock(spec=Font)
    # Deliberately don't set all_caps, small_caps etc.
    src.font.bold = True

    tgt.font = Mock(spec=Font)

    # Should not raise an error
    painter._copy_run_format(src, tgt)

    # Basic attributes should still be copied
    assert tgt.font.bold is True


def test_copy_font_error_handling_for_color():
    """Test that color copy errors are caught and logged."""
    painter = FormatPainter()
    src = Mock(spec=Run)
    tgt = Mock(spec=Run)

    # Source with color
    src.font = Mock(spec=Font)
    src.font.bold = True

    # Mock color that raises error when setting RGB
    mock_color = Mock()
    mock_color.rgb = (255, 0, 0)

    src.font.color = mock_color

    tgt.font = Mock(spec=Font)

    # Mock target color that raises error
    tgt.font.color = Mock()
    tgt.font.color.rgb = Mock(side_effect=Exception("Color not supported"))

    # Should not raise - error is caught and logged
    painter._copy_run_format(src, tgt)

    # Other attributes should still be copied
    assert tgt.font.bold is True


def test_copy_font_with_highlight_color():
    """Test copying highlight_color attribute."""
    painter = FormatPainter()
    src = Mock(spec=Run)
    tgt = Mock(spec=Run)

    # Source with highlight
    src.font = Mock(spec=Font)
    src.font.highlight_color = 7  # WD_COLOR_INDEX.YELLOW

    tgt.font = Mock(spec=Font)

    painter._copy_run_format(src, tgt)

    assert tgt.font.highlight_color == 7
