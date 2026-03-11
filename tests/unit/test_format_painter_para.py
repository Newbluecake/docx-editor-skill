import pytest
from unittest.mock import Mock
from docx.text.paragraph import Paragraph
from docx.text.parfmt import ParagraphFormat
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Pt, Inches
from docx_editor_skill.core.format_painter import FormatPainter

def test_copy_paragraph_format_basic():
    painter = FormatPainter()
    src = Mock(spec=Paragraph)
    tgt = Mock(spec=Paragraph)

    # Setup source format
    src_fmt = Mock(spec=ParagraphFormat)
    src_fmt.alignment = WD_ALIGN_PARAGRAPH.CENTER
    src_fmt.left_indent = Inches(1)
    src_fmt.right_indent = Inches(0.5)
    src_fmt.first_line_indent = Inches(0.25)
    src_fmt.space_before = Pt(12)
    src_fmt.space_after = Pt(6)
    src_fmt.line_spacing = 1.5
    src_fmt.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    src_fmt.keep_together = True
    src_fmt.keep_with_next = False
    src_fmt.page_break_before = True

    # Associate format with paragraph
    src.paragraph_format = src_fmt
    # alignment is also exposed directly on Paragraph often, but usually proxied to paragraph_format
    # python-docx Paragraph.alignment delegates to paragraph_format.alignment
    # We should test that we copy via paragraph_format mostly

    # Setup target
    tgt.paragraph_format = Mock(spec=ParagraphFormat)

    painter._copy_paragraph_format(src, tgt)

    # Assertions
    assert tgt.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert tgt.paragraph_format.left_indent == Inches(1)
    assert tgt.paragraph_format.right_indent == Inches(0.5)
    assert tgt.paragraph_format.first_line_indent == Inches(0.25)
    assert tgt.paragraph_format.space_before == Pt(12)
    assert tgt.paragraph_format.space_after == Pt(6)
    assert tgt.paragraph_format.line_spacing == 1.5
    assert tgt.paragraph_format.line_spacing_rule == WD_LINE_SPACING.ONE_POINT_FIVE
    assert tgt.paragraph_format.keep_together is True
    assert tgt.paragraph_format.keep_with_next is False
    assert tgt.paragraph_format.page_break_before is True

def test_copy_paragraph_style():
    """Test that style name is copied if possible"""
    painter = FormatPainter()
    src = Mock(spec=Paragraph)
    tgt = Mock(spec=Paragraph)

    src.style = "Heading 1"

    # Target mock needs to accept style assignment
    tgt.style = "Normal"

    # Mock paragraph_format to avoid errors in the method
    src.paragraph_format = Mock(spec=ParagraphFormat)
    tgt.paragraph_format = Mock(spec=ParagraphFormat)

    painter._copy_paragraph_format(src, tgt)

    assert tgt.style == "Heading 1"
