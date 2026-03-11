from dataclasses import dataclass
from typing import Optional, Literal

# Valid position types for the cursor
PositionType = Literal["before", "after", "inside_start", "inside_end"]

@dataclass
class Cursor:
    """
    Represents a specific insertion point within the document structure.

    The cursor defines where the next content insertion should happen.
    It can be positioned relative to an existing element (before/after)
    or relative to a container (inside_start/inside_end).

    Attributes:
        parent_id (str): ID of the container element. Defaults to "document_body" for the main document.
        element_id (Optional[str]): ID of the reference element (paragraph, table, run).
                                    If None, implies start/end of parent container.
        position (PositionType): Relation to the element/container.
                                 "before": Before the reference element.
                                 "after": After the reference element.
                                 "inside_start": At the beginning of the container.
                                 "inside_end": At the end of the container.
    """
    parent_id: str = "document_body"
    element_id: Optional[str] = None
    position: PositionType = "inside_end"

    def is_valid(self) -> bool:
        """Check if cursor state is valid."""
        valid_positions = ["before", "after", "inside_start", "inside_end"]
        if self.position not in valid_positions:
            return False

        # If position is before/after, we must have an element_id
        if self.position in ["before", "after"] and not self.element_id:
            return False

        return True
