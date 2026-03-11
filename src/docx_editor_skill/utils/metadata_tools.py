"""Metadata generation utilities for tracking element operations."""
import time
from typing import Dict, Any, Optional


class MetadataTools:
    """
    Utilities for creating consistent metadata across document operations.

    Standardizes metadata generation for create/copy/modify operations to ensure
    consistent tracking of element lineage and operation history.
    """

    @staticmethod
    def create_copy_metadata(
        source_id: Optional[str] = None,
        source_type: Optional[str] = None,
        operation_type: str = "copy",
        **additional_fields
    ) -> Dict[str, Any]:
        """
        Create metadata for copy operations.

        Generates standardized metadata for tracking copied elements, including
        source information and timestamp.

        Args:
            source_id (str, optional): ID of the source element that was copied.
                If None, source tracking is omitted (e.g., for range copies).
            source_type (str, optional): Type of source element ("paragraph", "table", etc.).
                Should be provided if source_id is provided.
            operation_type (str): Type of operation ("copy", "range_copy", "duplicate").
                Defaults to "copy".
            **additional_fields: Additional metadata fields to include.

        Returns:
            Dict containing standardized copy metadata with the following guaranteed fields:
                - copied_at (float): Unix timestamp of the copy operation
                - operation_type (str): Type of copy operation

            And optionally (if source_id provided):
                - source_id (str): ID of source element
                - source_type (str): Type of source element

        Examples:
            >>> # Copy with full source tracking
            >>> meta = MetadataTools.create_copy_metadata(
            ...     source_id="para_123",
            ...     source_type="paragraph"
            ... )
            >>> # Result: {"copied_at": 1234567890.0, "operation_type": "copy",
            ...            "source_id": "para_123", "source_type": "paragraph"}

            >>> # Range copy without specific source
            >>> meta = MetadataTools.create_copy_metadata(
            ...     operation_type="range_copy",
            ...     range_start="para_100",
            ...     range_end="para_150"
            ... )
            >>> # Result: {"copied_at": 1234567890.0, "operation_type": "range_copy",
            ...            "range_start": "para_100", "range_end": "para_150"}
        """
        metadata = {
            "copied_at": time.time(),
            "operation_type": operation_type
        }

        # Add source tracking if provided
        if source_id is not None:
            metadata["source_id"] = source_id
            if source_type is not None:
                metadata["source_type"] = source_type

        # Add any additional fields
        metadata.update(additional_fields)

        return metadata

    @staticmethod
    def create_creation_metadata(
        created_by: Optional[str] = None,
        **additional_fields
    ) -> Dict[str, Any]:
        """
        Create metadata for element creation operations.

        Generates standardized metadata for tracking newly created elements.

        Args:
            created_by (str, optional): Function or operation that created the element.
            **additional_fields: Additional metadata fields to include.

        Returns:
            Dict containing standardized creation metadata with the following guaranteed fields:
                - created_at (float): Unix timestamp of the creation

        Examples:
            >>> meta = MetadataTools.create_creation_metadata(created_by="docx_insert_paragraph")
            >>> # Result: {"created_at": 1234567890.0, "created_by": "docx_insert_paragraph"}
        """
        metadata = {
            "created_at": time.time()
        }

        if created_by is not None:
            metadata["created_by"] = created_by

        # Add any additional fields
        metadata.update(additional_fields)

        return metadata

    @staticmethod
    def create_modification_metadata(
        modified_by: Optional[str] = None,
        **additional_fields
    ) -> Dict[str, Any]:
        """
        Create metadata for element modification operations.

        Generates standardized metadata for tracking element modifications.

        Args:
            modified_by (str, optional): Function or operation that modified the element.
            **additional_fields: Additional metadata fields to include.

        Returns:
            Dict containing standardized modification metadata with the following guaranteed fields:
                - modified_at (float): Unix timestamp of the modification

        Examples:
            >>> meta = MetadataTools.create_modification_metadata(
            ...     modified_by="docx_update_paragraph_text",
            ...     old_length=100,
            ...     new_length=150
            ... )
        """
        metadata = {
            "modified_at": time.time()
        }

        if modified_by is not None:
            metadata["modified_by"] = modified_by

        # Add any additional fields
        metadata.update(additional_fields)

        return metadata
