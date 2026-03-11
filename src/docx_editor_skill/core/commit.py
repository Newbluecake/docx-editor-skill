"""
Commit data structure for change tracking.

This module provides the Commit class for tracking document changes.
"""

from dataclasses import dataclass
from typing import Dict, Any, List, Optional
from datetime import datetime, timezone
import uuid


@dataclass
class Commit:
    """
    Represents a single commit in the document history.

    A commit captures a snapshot of changes made to document elements,
    including before/after states and context information.

    Attributes:
        commit_id: Unique identifier for this commit (UUID)
        timestamp: ISO 8601 formatted timestamp
        operation: Name of the operation (e.g., "update_paragraph_text")
        changes: Dictionary containing before/after states and context
        affected_elements: List of element IDs affected by this commit
        description: Optional human-readable description
        user_metadata: Optional user-defined metadata
    """

    commit_id: str
    timestamp: str
    operation: str
    changes: Dict[str, Any]
    affected_elements: List[str]
    description: str = ""
    user_metadata: Optional[Dict[str, Any]] = None

    @classmethod
    def create(
        cls,
        operation: str,
        changes: Dict[str, Any],
        affected_elements: List[str],
        description: str = ""
    ) -> "Commit":
        """
        Create a new commit with auto-generated ID and timestamp.

        Args:
            operation: Name of the operation
            changes: Dictionary with 'before', 'after', and 'context' keys
            affected_elements: List of element IDs
            description: Optional description

        Returns:
            New Commit instance
        """
        return cls(
            commit_id=str(uuid.uuid4()),
            timestamp=datetime.now(timezone.utc).isoformat(),
            operation=operation,
            changes=changes,
            affected_elements=affected_elements,
            description=description
        )

    def to_dict(self) -> Dict[str, Any]:
        """
        Convert commit to dictionary for serialization.

        Returns:
            Dictionary representation of the commit
        """
        return {
            "commit_id": self.commit_id,
            "timestamp": self.timestamp,
            "operation": self.operation,
            "changes": self.changes,
            "affected_elements": self.affected_elements,
            "description": self.description,
            "user_metadata": self.user_metadata
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "Commit":
        """
        Create commit from dictionary.

        Args:
            data: Dictionary containing commit data

        Returns:
            Commit instance
        """
        return cls(
            commit_id=data["commit_id"],
            timestamp=data["timestamp"],
            operation=data["operation"],
            changes=data["changes"],
            affected_elements=data["affected_elements"],
            description=data.get("description", ""),
            user_metadata=data.get("user_metadata")
        )
