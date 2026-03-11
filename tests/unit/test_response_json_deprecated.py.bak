"""Unit tests for response formatting utilities."""

import json
import pytest
from docx_mcp_server.core.response import (
    ToolResponse,
    CursorInfo,
    create_success_response,
    create_error_response,
    create_context_aware_response
)


def test_cursor_info_to_dict():
    """Test CursorInfo serialization."""
    cursor = CursorInfo(
        element_id="para_123",
        position="after",
        parent_id="cell_456",
        context="Cursor: after Paragraph para_123"
    )

    result = cursor.to_dict()
    assert result["element_id"] == "para_123"
    assert result["position"] == "after"
    assert result["parent_id"] == "cell_456"
    assert "context" in result


def test_cursor_info_excludes_none():
    """Test that None values are excluded from CursorInfo dict."""
    cursor = CursorInfo(element_id="para_123")
    result = cursor.to_dict()

    assert "element_id" in result
    assert "position" in result  # Has default value
    assert "parent_id" not in result  # None should be excluded
    assert "context" not in result  # None should be excluded


def test_tool_response_to_json():
    """Test ToolResponse JSON serialization."""
    response = ToolResponse(
        status="success",
        message="Operation completed",
        data={"element_id": "para_123"}
    )

    json_str = response.to_json()
    parsed = json.loads(json_str)

    assert parsed["status"] == "success"
    assert parsed["message"] == "Operation completed"
    assert parsed["data"]["element_id"] == "para_123"


def test_create_success_response_minimal():
    """Test creating minimal success response."""
    result = create_success_response("Paragraph created")
    parsed = json.loads(result)

    assert parsed["status"] == "success"
    assert parsed["message"] == "Paragraph created"
    assert parsed["data"] == {}


def test_create_success_response_with_element_id():
    """Test success response with element ID."""
    result = create_success_response(
        "Paragraph created",
        element_id="para_abc123"
    )
    parsed = json.loads(result)

    assert parsed["status"] == "success"
    assert parsed["data"]["element_id"] == "para_abc123"


def test_create_success_response_with_cursor():
    """Test success response with cursor information."""
    cursor = CursorInfo(
        element_id="para_123",
        position="after",
        context="Cursor: after para_123"
    )

    result = create_success_response(
        "Paragraph created",
        element_id="para_123",
        cursor=cursor
    )
    parsed = json.loads(result)

    assert parsed["status"] == "success"
    assert parsed["data"]["element_id"] == "para_123"
    assert "cursor" in parsed["data"]
    assert parsed["data"]["cursor"]["element_id"] == "para_123"
    assert parsed["data"]["cursor"]["position"] == "after"


def test_create_success_response_with_extra_data():
    """Test success response with additional data fields."""
    result = create_success_response(
        "Table created",
        element_id="table_123",
        rows=3,
        cols=2,
        style="TableGrid"
    )
    parsed = json.loads(result)

    assert parsed["data"]["element_id"] == "table_123"
    assert parsed["data"]["rows"] == 3
    assert parsed["data"]["cols"] == 2
    assert parsed["data"]["style"] == "TableGrid"


def test_create_error_response():
    """Test error response creation."""
    result = create_error_response("Session not found")
    parsed = json.loads(result)

    assert parsed["status"] == "error"
    assert parsed["message"] == "Session not found"


def test_create_error_response_with_type():
    """Test error response with error type."""
    result = create_error_response(
        "Invalid element ID",
        error_type="ValidationError"
    )
    parsed = json.loads(result)

    assert parsed["status"] == "error"
    assert parsed["data"]["error_type"] == "ValidationError"


def test_create_context_aware_response_without_cursor():
    """Test context-aware response when cursor is disabled."""
    # Mock session without cursor
    class MockSession:
        pass

    session = MockSession()
    result = create_context_aware_response(
        session,
        "Paragraph created",
        element_id="para_123",
        include_cursor=False
    )
    parsed = json.loads(result)

    assert parsed["status"] == "success"
    assert parsed["data"]["element_id"] == "para_123"
    assert "cursor" not in parsed["data"]


def test_create_context_aware_response_with_cursor():
    """Test context-aware response with cursor context."""
    # Mock session with cursor
    class MockCursor:
        element_id = "para_123"
        position = "after"
        parent_id = None

    class MockSession:
        cursor = MockCursor()

        def get_cursor_context(self):
            return "Cursor: after Paragraph para_123"

    session = MockSession()
    result = create_context_aware_response(
        session,
        "Paragraph created",
        element_id="para_123"
    )
    parsed = json.loads(result)

    assert parsed["status"] == "success"
    assert "cursor" in parsed["data"]
    assert parsed["data"]["cursor"]["element_id"] == "para_123"
    assert "context" in parsed["data"]["cursor"]


def test_json_serialization_with_special_characters():
    """Test JSON serialization handles special characters correctly."""
    result = create_success_response(
        "Created paragraph with text: \"Hello\\nWorld\"",
        element_id="para_123"
    )
    parsed = json.loads(result)

    assert "Hello\\nWorld" in parsed["message"]
    assert parsed["status"] == "success"


def test_response_serialization_failure_fallback():
    """Test that serialization failures are handled gracefully."""
    # Create a response with non-serializable data
    response = ToolResponse(
        status="success",
        message="Test",
        data={"func": lambda x: x}  # Functions are not JSON serializable
    )

    json_str = response.to_json()
    parsed = json.loads(json_str)

    # Should fallback to error response
    assert parsed["status"] == "error"
    assert "serialization failed" in parsed["message"].lower()
