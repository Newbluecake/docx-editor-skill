import pytest
from unittest.mock import MagicMock, Mock
from docx.table import Table
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree
from docx_editor_skill.core.format_painter import FormatPainter

# Helper to create a mock table with OXML elements
def create_mock_table():
    table = Mock(spec=Table)

    # Create a real CT_Tbl object so property access works
    tbl = OxmlElement('w:tbl')
    tblPr = OxmlElement('w:tblPr')
    tbl.append(tblPr)

    # Attach _element
    table._element = tbl

    # Mock style property
    table.style = "Table Grid"
    return table

def test_copy_table_format_style():
    painter = FormatPainter()
    src = create_mock_table()
    tgt = create_mock_table()

    src.style = "List Table 1 Light"
    tgt.style = "Table Grid"

    painter._copy_table_format(src, tgt)

    assert tgt.style == "List Table 1 Light"

def test_copy_table_borders_xml():
    painter = FormatPainter()
    src = create_mock_table()
    tgt = create_mock_table()

    # Add borders to source
    src_tblPr = src._element.tblPr
    borders = OxmlElement('w:tblBorders')
    top = OxmlElement('w:top')
    top.set(qn('w:val'), 'single')
    borders.append(top)
    src_tblPr.append(borders)

    # Ensure target has no borders initially
    tgt_tblPr = tgt._element.tblPr
    assert tgt_tblPr.find(qn('w:tblBorders')) is None

    painter._copy_table_format(src, tgt)

    # Verify target has borders copied
    tgt_borders = tgt_tblPr.find(qn('w:tblBorders'))
    assert tgt_borders is not None
    tgt_top = tgt_borders.find(qn('w:top'))
    assert tgt_top is not None
    assert tgt_top.get(qn('w:val')) == 'single'

def test_copy_table_shading_xml():
    painter = FormatPainter()
    src = create_mock_table()
    tgt = create_mock_table()

    # Add shading to source
    src_tblPr = src._element.tblPr
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), 'FF0000')
    src_tblPr.append(shd)

    painter._copy_table_format(src, tgt)

    # Verify target has shading copied
    tgt_shd = tgt._element.tblPr.find(qn('w:shd'))
    assert tgt_shd is not None
    assert tgt_shd.get(qn('w:fill')) == 'FF0000'
