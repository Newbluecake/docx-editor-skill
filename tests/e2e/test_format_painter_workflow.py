import pytest
from tests.helpers import (
    extract_session_id,
    extract_element_id,
    extract_metadata_field,
    extract_all_metadata,
    is_success,
    is_error
)
from tests.helpers.session_helpers import setup_active_session, teardown_active_session
import os
import json
from docx_editor_skill.tools.session_tools import docx_save
from docx_editor_skill.tools.paragraph_tools import docx_insert_paragraph
from docx_editor_skill.tools.run_tools import docx_insert_run, docx_set_font
from docx_editor_skill.tools.format_tools import docx_set_alignment, docx_format_copy
from docx_editor_skill.tools.table_tools import docx_insert_table, docx_get_cell
from docx_editor_skill.core.session import session_manager
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def _extract(response):
    if not is_success(response):
        raise ValueError(f"Tool failed: {response}")
    return extract_all_metadata(response)

@pytest.fixture
def session_id():
    # Setup
    sid = setup_active_session()
    yield sid
    # Teardown
    teardown_active_session()
    if os.path.exists("test_format_painter_e2e.docx"):
        os.remove("test_format_painter_e2e.docx")

def test_format_painter_workflow(session_id):
    """
    E2E Test for Format Painter Workflow:
    1. Create Source Elements (Para, Run) with specific formatting.
    2. Create Target Elements (Plain).
    3. Use docx_format_copy to transfer format.
    4. Verify persistence and application.
    """
    # 1. Setup Source Paragraph (Centered)
    src_para_resp = docx_insert_paragraph("Source Paragraph", position="end:document_body")
    src_para_id = _extract(src_para_resp)["element_id"]
    docx_set_alignment(src_para_id, "center")

    # 2. Setup Source Run (Bold, Red, Size 16)
    src_run_resp = docx_insert_run("Source Run", position=f"inside:{src_para_id}")
    src_run_id = _extract(src_run_resp)["element_id"]
    docx_set_font(src_run_id, bold=True, size=16, color_hex="FF0000")

    # 3. Setup Target Paragraph (Left aligned by default)
    tgt_para_resp = docx_insert_paragraph("Target Paragraph", position="end:document_body")
    tgt_para_id = _extract(tgt_para_resp)["element_id"]

    # 4. Setup Target Run (Plain)
    tgt_run_resp = docx_insert_run("Target Run", position=f"inside:{tgt_para_id}")
    tgt_run_id = _extract(tgt_run_resp)["element_id"]

    # 5. Apply Format Painter: Para -> Para
    docx_format_copy(src_para_id, tgt_para_id)

    # 6. Apply Format Painter: Run -> Run
    docx_format_copy(src_run_id, tgt_run_id)

    # 7. Verification
    session = session_manager.get_session(session_id)
    tgt_para = session.get_object(tgt_para_id)
    tgt_run = session.get_object(tgt_run_id)

    # Verify Paragraph alignment
    assert tgt_para.alignment == WD_ALIGN_PARAGRAPH.CENTER

    # Verify Run font
    assert tgt_run.font.bold is True
    assert tgt_run.font.size == Pt(16)
    assert tgt_run.font.color.rgb == RGBColor(255, 0, 0)

    # 8. Test Cross-Type: Para -> Run (Smart Matching)
    heading_resp = docx_insert_paragraph("Heading", position="end:document_body", style="Heading 1")
    heading_id = _extract(heading_resp)["element_id"]

    plain_run_resp = docx_insert_run("Plain Run", position=f"inside:{tgt_para_id}")
    plain_run_id = _extract(plain_run_resp)["element_id"]

    docx_format_copy(heading_id, plain_run_id)

    plain_run = session.get_object(plain_run_id)

def test_format_painter_table_workflow(session_id):
    """Test table format copying workflow"""
    # Create Source Table
    src_table_resp = docx_insert_table(2, 2, position="end:document_body")
    src_table_id = _extract(src_table_resp)["element_id"]

    session = session_manager.get_session(session_id)
    src_table = session.get_object(src_table_id)
    src_table.style = "Table Grid"

    # Create Target Table
    tgt_table_resp = docx_insert_table(2, 2, position="end:document_body")
    tgt_table_id = _extract(tgt_table_resp)["element_id"]

    tgt_table = session.get_object(tgt_table_id)
    tgt_table.style = "Normal Table" # or something else

    # Copy
    docx_format_copy(src_table_id, tgt_table_id)

    # Verify
    assert tgt_table.style.name == "Table Grid"
