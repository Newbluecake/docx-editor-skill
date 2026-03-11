import pytest
from unittest.mock import MagicMock
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx_editor_skill.core.properties import set_properties

def test_set_font_properties():
    mock_obj = MagicMock()
    mock_font = MagicMock()
    mock_obj.font = mock_font

    props = {
        "font": {
            "name": "Arial",
            "size": 12,
            "bold": True,
            "color": "FF0000"
        }
    }

    set_properties(mock_obj, props)

    assert mock_font.name == "Arial"
    assert mock_font.size == Pt(12)
    assert mock_font.bold is True
    # RGB check is tricky with mocks if logic does assignment, but let's assume it tried
    # We can check if color.rgb was accessed
    assert mock_font.color.rgb == RGBColor(255, 0, 0)

def test_set_paragraph_properties():
    mock_obj = MagicMock()
    mock_fmt = MagicMock()
    mock_obj.paragraph_format = mock_fmt

    props = {
        "paragraph_format": {
            "alignment": "center",
            "space_after": 12
        }
    }

    set_properties(mock_obj, props)

    assert mock_fmt.alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert mock_fmt.space_after == Pt(12)

def test_set_direct_alignment():
    mock_obj = MagicMock()
    # Mock paragraph_format as well since logic tries to set it there if alignment attr missing
    mock_obj.paragraph_format = MagicMock()

    # Case 1: object has .alignment
    mock_obj.alignment = None
    set_properties(mock_obj, {"alignment": "right"})
    assert mock_obj.alignment == WD_ALIGN_PARAGRAPH.RIGHT

def test_set_vertical_alignment():
    mock_obj = MagicMock()
    mock_obj.vertical_alignment = None

    set_properties(mock_obj, {"vertical_alignment": "center"})
    assert mock_obj.vertical_alignment == WD_ALIGN_VERTICAL.CENTER
