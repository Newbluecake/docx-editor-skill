"""
Unit tests for Commit data structure.
"""

import pytest
from datetime import datetime
from docx_editor_skill.core.commit import Commit


def test_commit_creation():
    """Test basic commit creation."""
    commit = Commit(
        commit_id="test-id-123",
        timestamp="2026-01-22T10:00:00Z",
        operation="update_paragraph_text",
        changes={
            "before": {"text": "old text"},
            "after": {"text": "new text"},
            "context": {"parent_id": "document_body"}
        },
        affected_elements=["para_123"]
    )

    assert commit.commit_id == "test-id-123"
    assert commit.timestamp == "2026-01-22T10:00:00Z"
    assert commit.operation == "update_paragraph_text"
    assert commit.changes["before"]["text"] == "old text"
    assert commit.changes["after"]["text"] == "new text"
    assert commit.affected_elements == ["para_123"]
    assert commit.description == ""


def test_commit_create_factory():
    """Test commit creation using factory method."""
    commit = Commit.create(
        operation="set_font",
        changes={
            "before": {"bold": False},
            "after": {"bold": True}
        },
        affected_elements=["run_456"],
        description="Make text bold"
    )

    assert commit.commit_id is not None
    assert len(commit.commit_id) == 36  # UUID format
    assert "+00:00" in commit.timestamp or commit.timestamp.endswith("Z")  # UTC timezone
    assert commit.operation == "set_font"
    assert commit.description == "Make text bold"


def test_commit_to_dict():
    """Test commit serialization to dictionary."""
    commit = Commit(
        commit_id="test-id",
        timestamp="2026-01-22T10:00:00Z",
        operation="test_op",
        changes={"before": {}, "after": {}},
        affected_elements=["elem_1"],
        description="Test commit"
    )

    data = commit.to_dict()

    assert data["commit_id"] == "test-id"
    assert data["timestamp"] == "2026-01-22T10:00:00Z"
    assert data["operation"] == "test_op"
    assert data["description"] == "Test commit"
    assert data["affected_elements"] == ["elem_1"]


def test_commit_from_dict():
    """Test commit deserialization from dictionary."""
    data = {
        "commit_id": "test-id",
        "timestamp": "2026-01-22T10:00:00Z",
        "operation": "test_op",
        "changes": {"before": {"text": "old"}, "after": {"text": "new"}},
        "affected_elements": ["elem_1", "elem_2"],
        "description": "Test commit"
    }

    commit = Commit.from_dict(data)

    assert commit.commit_id == "test-id"
    assert commit.operation == "test_op"
    assert commit.changes["before"]["text"] == "old"
    assert commit.affected_elements == ["elem_1", "elem_2"]


def test_commit_with_metadata():
    """Test commit with user metadata."""
    commit = Commit.create(
        operation="test_op",
        changes={"before": {}, "after": {}},
        affected_elements=["elem_1"]
    )
    commit.user_metadata = {"author": "test_user", "tags": ["important"]}

    data = commit.to_dict()
    assert data["user_metadata"]["author"] == "test_user"

    restored = Commit.from_dict(data)
    assert restored.user_metadata["tags"] == ["important"]
