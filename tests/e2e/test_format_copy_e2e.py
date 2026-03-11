import pytest
import ast
from tests.helpers import (
    extract_session_id,
    extract_element_id,
    extract_metadata_field,
    extract_all_metadata,
    is_success,
    is_error,
    extract_element_ids_list
)
from tests.helpers.session_helpers import setup_active_session, teardown_active_session
import json
import os
from docx_editor_skill.tools.session_tools import docx_save
from docx_editor_skill.tools.paragraph_tools import docx_insert_heading, docx_insert_paragraph, docx_copy_paragraph
from docx_editor_skill.tools.run_tools import docx_insert_run, docx_set_font
from docx_editor_skill.tools.copy_tools import docx_copy_elements_range
from docx_editor_skill.tools.format_tools import docx_extract_format_template, docx_apply_format_template
from docx_editor_skill.tools.advanced_tools import docx_batch_replace_text
from docx_editor_skill.tools.content_tools import docx_read_content

def _extract(response):
    if not is_success(response):
        raise ValueError(f"Tool failed: {response}")

    metadata = extract_all_metadata(response)

    # Handle template field which might be a JSON string or dict string
    if 'template' in metadata and isinstance(metadata['template'], str):
        try:
            # Try JSON first
            metadata['template'] = json.loads(metadata['template'])
        except json.JSONDecodeError:
            try:
                metadata['template'] = ast.literal_eval(metadata['template'])
            except (ValueError, SyntaxError):
                pass

    return metadata

def test_format_copy_system_workflow(tmp_path):
    """
    E2E scenario:
    1. Create a "Template" section with specific styling.
    2. Extract formatting from the template.
    3. Copy the template section (Range Copy).
    4. Apply extracted formatting to a new raw paragraph.
    5. Batch replace placeholders in the copied section.
    """
    # 1. Setup
    output_file = tmp_path / "e2e_format_copy.docx"
    session_id = setup_active_session()
    # --- Step 1: Create Template Content ---
    # Heading 1
    h1_resp = docx_insert_heading("Template Section", position="end:document_body", level=1)
    h1_id = _extract(h1_resp)["element_id"]

    # Styled Paragraph
    p_resp = docx_insert_paragraph("", position="end:document_body")
    p_id = _extract(p_resp)["element_id"]

    run_resp = docx_insert_run("Key: ", position=f"inside:{p_id}")
    run_id = _extract(run_resp)["element_id"]
    docx_set_font(run_id, bold=True, color_hex="0000FF") # Blue bold label

    val_run_resp = docx_insert_run("{VALUE}", position=f"inside:{p_id}")
    val_run_id = _extract(val_run_resp)["element_id"]
    docx_set_font(val_run_id, italic=True) # Italic placeholder

    # End marker for range
    end_p_resp = docx_insert_paragraph("--- End Template ---", position="end:document_body")
    end_p_id = _extract(end_p_resp)["element_id"]

    # --- Step 2: Extract Formatting Template ---
    # Extract the blue bold style from the first run
    template_resp = docx_extract_format_template(run_id)
    template_data = _extract(template_resp)
    template_json = json.dumps(template_data["template"])

    assert "0000FF" in template_json
    assert "bold" in template_json

    # --- Step 3: Range Copy ---
    # Copy from Heading to End Marker
    # We expect 3 elements: Heading, Styled Para, End Marker
    new_ids_resp = docx_copy_elements_range(h1_id, end_p_id, position="end:document_body")

    # Extract IDs from Markdown response
    new_ids = extract_element_ids_list(new_ids_resp)

    assert len(new_ids) == 3

    # --- Step 4: Apply Formatting Template ---
    # Create a raw paragraph
    raw_p_resp = docx_insert_paragraph("", position="end:document_body")
    raw_p_id = _extract(raw_p_resp)["element_id"]

    raw_run_resp = docx_insert_run("Applied Style Text", position=f"inside:{raw_p_id}")
    raw_run_id = _extract(raw_run_resp)["element_id"]

    # Apply the blue bold template
    docx_apply_format_template(raw_run_id, template_json)

    # --- Step 5: Batch Replace ---
    # Replace {VALUE} with "Actual Data" globally (should affect original and copy)
    replacements = {
        "{VALUE}": "Actual Data",
        "Template Section": "Generated Section"
    }
    docx_batch_replace_text(json.dumps(replacements))

    # --- Verification ---
    content_resp = docx_read_content()
    content = content_resp  # This tool returns plain text

    print(f"Document Content:\n{content}")

    # Check replacements
    assert "Key: Actual Data" in content # Replaced value
    assert "Generated Section" in content # Replaced heading

    # Check that we have the copied section (it appears twice now, effectively)
    # Original: Generated Section ... Key: Actual Data ... --- End Template ---
    # Copy: Generated Section ... Key: Actual Data ... --- End Template ---
    assert content.count("Generated Section") >= 2
    assert content.count("Key: Actual Data") >= 2

    # Save
    docx_save(str(output_file))
    teardown_active_session()

    assert os.path.exists(output_file)
