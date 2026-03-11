"""
Unit tests for Response formatting with change tracking.
"""

import pytest
import json
from docx import Document
from docx_editor_skill.core.session import Session
from docx_editor_skill.core.response import (
    create_change_tracked_response,
    create_success_response
)

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


def test_create_change_tracked_response():
    """Test creating response with change tracking."""
    session = Session(session_id="test", document=Document())

    changes = {
        "before": {"text": "old text"},
        "after": {"text": "new text"},
        "context": {"parent_id": "document_body"}
    }

    response = create_change_tracked_response(
        session=session,
        message="Updated successfully",
        element_id="para_123",
        changes=changes,
        commit_id="commit-456",
        include_cursor=False
    )

    data = json.loads(response)

    assert is_success(response)
    assert extract_metadata_field(response, "element_id") == "para_123"
    assert "old text" in response
    assert extract_metadata_field(response, "commit_id") == "commit-456"


def test_change_tracked_response_with_extra_data():
    """Test change tracked response with additional data fields."""
    session = Session(session_id="test", document=Document())

    response = create_change_tracked_response(
        session=session,
        message="Operation completed",
        element_id="elem_789",
        changes={"before": {}, "after": {}},
        commit_id="commit-abc",
        include_cursor=False,
        custom_field="custom_value",
        another_field=123
    )

    data = json.loads(response)

    assert is_success(response)
    assert extract_metadata_field(response, "custom_field") == "custom_value"
    assert extract_metadata_field(response, "another_field") == 123


def test_backward_compatibility():
    """Test that existing create_success_response still works."""
    response = create_success_response(
        message="Success",
        element_id="elem_123",
        custom_data="value"
    )

    data = json.loads(response)

    assert is_success(response)
    assert extract_metadata_field(response, "element_id") == "elem_123"
    assert extract_metadata_field(response, "custom_data") == "value"
