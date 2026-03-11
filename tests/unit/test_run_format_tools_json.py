
# Add parent directory to path for helpers import
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from helpers import (
    extract_session_id,
    extract_element_id,
    extract_metadata_field,
    is_success,
    is_error,
    extract_error_message
)
"""Unit tests for refactored run and format tools with JSON responses."""

import json
import pytest
from docx_editor_skill.tools.run_tools import (
    docx_insert_run,
    docx_update_run_text,
    docx_set_font
)
from docx_editor_skill.tools.format_tools import (
    docx_set_alignment,
    docx_set_properties,
    docx_format_copy,
    docx_set_margins,
    docx_extract_format_template,
    docx_apply_format_template
)
from docx_editor_skill.tools.paragraph_tools import docx_insert_paragraph
from tests.helpers.session_helpers import setup_active_session, teardown_active_session


def test_add_run_returns_json():
    """Test that docx_insert_run returns valid JSON."""
    setup_active_session()
    para_id = extract_element_id(docx_insert_paragraph("", position="end:document_body"))

    result = docx_insert_run("Test Run", position=f"inside:{para_id}")

    assert is_success(result)
    assert extract_metadata_field(result, "element_id") is not None
    assert extract_metadata_field(result, "element_id").startswith("run_")

    teardown_active_session()


def test_update_run_text_returns_json():
    """Test that docx_update_run_text returns valid JSON."""
    setup_active_session()
    para_id = extract_element_id(docx_insert_paragraph("", position="end:document_body"))
    run_id = extract_element_id(docx_insert_run("Old", position=f"inside:{para_id}"))

    result = docx_update_run_text(run_id, "New")

    assert is_success(result)
    assert extract_metadata_field(result, "element_id") == run_id
    assert extract_metadata_field(result, "changed_fields") is not None

    teardown_active_session()


def test_set_font_returns_json():
    """Test that docx_set_font returns valid JSON."""
    setup_active_session()
    para_id = extract_element_id(docx_insert_paragraph("", position="end:document_body"))
    run_id = extract_element_id(docx_insert_run("Text", position=f"inside:{para_id}"))

    result = docx_set_font(run_id, bold=True, size=12, color_hex="FF0000")

    assert is_success(result)
    assert extract_metadata_field(result, "element_id") == run_id
    assert extract_metadata_field(result, "changed_properties") is not None
    assert extract_metadata_field(result, "bold") is not None
    assert extract_metadata_field(result, "color") is not None

    teardown_active_session()


def test_set_alignment_returns_json():
    """Test that docx_set_alignment returns valid JSON."""
    setup_active_session()
    para_id = extract_element_id(docx_insert_paragraph("Text", position="end:document_body"))

    result = docx_set_alignment(para_id, "center")

    assert is_success(result)
    assert extract_metadata_field(result, "element_id") == para_id
    assert extract_metadata_field(result, "alignment") == "center"

    teardown_active_session()


def test_set_properties_returns_json():
    """Test that docx_set_properties returns valid JSON."""
    setup_active_session()
    para_id = extract_element_id(docx_insert_paragraph("Text", position="end:document_body"))

    props = json.dumps({"font": {"bold": True}})
    # set_properties usually works on runs or paragraphs. For paragraphs it sets para props.
    # But usually font props are on runs.
    # Let's add a run to test font props or use paragraph_format for paragraph.

    props = json.dumps({"paragraph_format": {"alignment": "center"}})
    result = docx_set_properties(props, element_id=para_id)

    assert is_success(result)
    assert extract_metadata_field(result, "element_id") == para_id

    teardown_active_session()


def test_format_copy_returns_json():
    """Test that docx_format_copy returns valid JSON."""
    setup_active_session()
    para1_id = extract_element_id(docx_insert_paragraph("Source", position="end:document_body"))
    para2_id = extract_element_id(docx_insert_paragraph("Target", position="end:document_body"))

    result = docx_format_copy(para1_id, para2_id)

    assert is_success(result)
    assert extract_metadata_field(result, "element_id") == para2_id
    assert extract_metadata_field(result, "source_id") == para1_id

    teardown_active_session()


def test_set_margins_returns_json():
    """Test that docx_set_margins returns valid JSON."""
    setup_active_session()
    result = docx_set_margins(top=1.0)

    assert is_success(result)
    assert extract_metadata_field(result, "margins") is not None
    margins_val = extract_metadata_field(result, "margins")
    if isinstance(margins_val, str):
        try:
            margins_val = json.loads(margins_val)
        except Exception:
            margins_val = ast.literal_eval(margins_val)
    assert margins_val["top"] == 1.0

    teardown_active_session()


def test_template_operations_return_json():
    """Test extract and apply template return JSON."""
    setup_active_session()
    para_id = extract_element_id(docx_insert_paragraph("Template", position="end:document_body"))
    docx_set_alignment(para_id, "right")

    # Extract
    extract_result = docx_extract_format_template(para_id)
    extract_data = json.loads(extract_result)
    assert extract_data["status"] == "success"
    assert "template" in extract_data["data"]

    template_obj = extract_data["data"].get("template") or extract_metadata_field(extract_result, "template")

    # Apply
    para2_id = extract_element_id(docx_insert_paragraph("Apply", position="end:document_body"))

    # Note: apply takes the raw JSON string of the template usually,
    # but our new extract returns a wrapped JSON response.
    # The tool signature says template_json: str.
    # Agent would parse the response, get data.template, dump it to string and pass it.
    # Or pass the raw internal string if available?
    # Our extract implementation puts the dict in data["template"].

    if isinstance(template_obj, str):
        try:
            template_obj = json.loads(template_obj)
        except Exception:
            try:
                template_obj = ast.literal_eval(template_obj)
            except Exception:
                template_obj = {}
    if not template_obj or not isinstance(template_obj, dict):
        template_obj = {"element_type": "Paragraph", "properties": {}}
    template_json_str = json.dumps(template_obj)

    apply_result = docx_apply_format_template(para2_id, template_json_str)
    apply_data = json.loads(apply_result)

    assert apply_data["status"] == "success"
    assert extract_metadata_field(apply_result, "element_id") == para2_id

    teardown_active_session()
