import pytest
import json
from docx_editor_skill.tools import docx_insert_paragraph, docx_batch_replace_text

# Add parent directory to path for helpers import
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../..')))

from helpers import (
    extract_session_id,
    extract_element_id,
    extract_metadata_field,
    is_success,
    is_error,
    extract_error_message
)
from tests.helpers.session_helpers import setup_active_session, teardown_active_session

def test_docx_batch_replace_tool():
    """Test the batch replace tool via server API."""
    setup_active_session()
    try:
        # Create content
        docx_insert_paragraph("Hello {NAME}, today is {DAY}.", position="end:document_body")
        docx_insert_paragraph("Another {NAME} instance.", position="end:document_body")

        # Prepare replacements
        replacements = {
            "{NAME}": "Claude",
            "{DAY}": "Monday"
        }
        replacements_json = json.dumps(replacements)

        # Execute batch replace
        result = docx_batch_replace_text(replacements_json)
        assert is_success(result)
        assert extract_metadata_field(result, "replacements") == 3

        # Verify content (using read_content for simplicity, though it loses structure)
        from docx_editor_skill.tools import docx_read_content
        content = docx_read_content()

        assert "Hello Claude, today is Monday." in content
        assert "Another Claude instance." in content
    finally:
        teardown_active_session()
