from typing import Any, Optional
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx_editor_skill.core.xml_util import ElementNavigator

class ContextVisualizer:
    """
    Generates ASCII tree visualizations of the document structure.
    Designed to be token-efficient and intuitive for LLMs.
    """

    def __init__(self, session=None):
        self.session = session

    def _get_element_id(self, element: Any) -> str:
        """Helper to get ID if available, otherwise '?'."""
        if not self.session:
            return "?"
        # Use internal protected method if available, or just check cache
        # The session usually has _get_element_id
        if hasattr(self.session, '_get_element_id'):
            return self.session._get_element_id(element, auto_register=False) or "?"
        return "?"

    def _get_element_summary(self, element: Any) -> str:
        """Get a concise summary of the element."""
        if isinstance(element, Paragraph):
            text = element.text.strip()
            # Truncate text
            if len(text) > 30:
                text = text[:27] + "..."
            elif not text:
                text = ""

            # Check style for Heading
            style_name = element.style.name if element.style else "Normal"
            if "Heading" in style_name:
                return f"{style_name}: \"{text}\""
            return f"Para: \"{text}\""

        elif isinstance(element, Table):
            rows = len(element.rows)
            cols = len(element.columns)
            return f"Table: {rows}x{cols}"

        return "Unknown Element"

    def generate_tree_view(self, target_element: Any, sibling_range: int = 1) -> str:
        """
        Generate a visual tree representation.

        Args:
            target_element: The element to focus on.
            sibling_range: Number of siblings to show before/after.
        """
        if not hasattr(target_element, '_element'):
            return "Element unavailable"

        # 1. Get Parent and Siblings
        # We iterate XML siblings because python-docx doesn't give us a robust sibling list directly
        parent = ElementNavigator.get_docx_parent(target_element, self.session.document if self.session else None)
        parent_name = "Container"

        children_xml = []
        if hasattr(parent, '_element'):
             children_xml = list(ElementNavigator.iter_children_xml(parent._element))
             if hasattr(parent, 'name'): # e.g. Section? No.
                 pass
             if isinstance(parent, Table): # Cell parent is Row
                 pass
             # Try to guess parent name
             if self.session and (parent == self.session.document or
                                  (hasattr(self.session.document, '_body') and parent == self.session.document._body)):
                 parent_name = "Document Body"
             elif hasattr(parent, 'text'):
                 parent_name = "Cell (Parent)"
        else:
            # Fallback if we can't resolve high-level parent:
            # Iterating direct XML siblings of target
            parent_xml = target_element._element.getparent()
            if parent_xml is not None:
                children_xml = list(ElementNavigator.iter_children_xml(parent_xml))
                parent_name = f"XML Container <{parent_xml.tag.split('}')[-1]}>"

        if not children_xml:
            return "Empty Container"

        # 2. Find index of target
        target_xml = target_element._element
        try:
            target_idx = children_xml.index(target_xml)
        except ValueError:
            return "Element detached from tree"

        # 3. Calculate range
        start_idx = max(0, target_idx - sibling_range)
        end_idx = min(len(children_xml), target_idx + sibling_range + 1)

        lines = [parent_name]

        # Add "..." if skipped previous siblings
        if start_idx > 0:
            lines.append("  ├── (...)")

        for i in range(start_idx, end_idx):
            child_xml = children_xml[i]
            is_target = (i == target_idx)

            # Convert XML back to docx object (temporarily) for summary
            # We assume simple wrapping via xml_util or just basic inference
            child_obj = None
            if child_xml.tag.endswith('p'):
                child_obj = Paragraph(child_xml, parent)
            elif child_xml.tag.endswith('tbl'):
                child_obj = Table(child_xml, parent)

            summary = self._get_element_summary(child_obj) if child_obj else f"<{child_xml.tag.split('}')[-1]}>"

            # Get ID
            # Note: We won't auto-register neighbors just for visualization to save memory/pollution
            # So neighbors might show ID "?" if not accessed before.
            # However, Requirements say "surrounding structure".
            # If we want IDs, we might need to register them or just show "?"
            # Let's show ID if available.
            eid = "?"
            if child_obj and self.session:
                eid = self.session._get_element_id(child_obj, auto_register=False) or "?"

            prefix = "└──" if (i == len(children_xml) - 1 and end_idx == len(children_xml)) else "├──"

            if is_target:
                # Highlight current
                lines.append(f"  {prefix} [{summary} ({eid})] <--- Current")
            else:
                lines.append(f"  {prefix} {summary} ({eid})")

        # Add "..." if skipped next siblings
        if end_idx < len(children_xml):
            lines.append("  └── (...)")

        return "\n".join(lines)


class PositionResolver:
    """
    Resolves position strings (e.g., "after:para_123") into actionable
    insertion targets (parent, ref_element, mode).
    """

    def __init__(self, session):
        self.session = session

    def resolve(self, position_str: Optional[str], default_parent=None):
        """
        Parse position string and resolve to document objects.

        Args:
            position_str: e.g. "after:para_123", "inside:table_1"
            default_parent: Fallback parent if position is None

        Returns:
            Tuple (parent_obj, ref_obj, mode)
            - parent_obj: The container to insert into
            - ref_obj: The anchor element (for before/after)
            - mode: "append", "before", "after", "start"
        """
        if not position_str:
            return default_parent or self.session.document, None, "append"

        parts = position_str.split(":", 1)
        if len(parts) != 2:
            raise ValueError(f"Invalid position format: '{position_str}'. Expected 'mode:id'")

        mode, target_id = parts[0].lower(), parts[1]

        if mode not in ["after", "before", "inside", "start", "end"]:
             raise ValueError(f"Invalid position mode: '{mode}'. Supported: after, before, inside, start, end")

        # Resolve special IDs first
        try:
            resolved_id = self.session.resolve_special_id(target_id)
        except ValueError as e:
            # Re-raise with position context
            raise ValueError(f"Position resolution failed: {str(e)}")

        # Get the target object using resolved ID
        target_obj = self.session.get_object(resolved_id)
        if not target_obj:
             # Try getting it directly if it's special ID like document_body
             if resolved_id == "document_body":
                 target_obj = self.session.document
             else:
                 raise ValueError(f"Target element '{target_id}' not found")

        # Resolve Logic
        parent = None
        ref = None

        if mode in ["after", "before"]:
            # Target is the sibling anchor
            ref = target_obj
            # Resolve parent
            parent = ElementNavigator.get_docx_parent(target_obj, self.session.document)
            if not parent:
                # Fallback: try to guess from session context or check if it's attached
                # If target has _element, we can get parent XML, but we need python-docx object for parent
                # Creating a temporary wrapper or failing?
                # For now, if we can't resolve high-level parent, we might fail unless we support XML-only insertion
                # But our tools usually expect a docx Parent object (e.g. body.add_paragraph).
                # Wait, we can use ElementManipulator for low-level XML insert,
                # but we still return a python-docx object wrapped with a parent.
                # So we really need the parent object.
                raise ValueError(f"Could not resolve parent for element '{target_id}'. Element might be detached.")

        elif mode in ["inside", "end"]:
            # Target IS the parent (container), append to end
            parent = target_obj
            mode = "append"

        elif mode == "start":
            # Target IS the parent, insert at beginning
            parent = target_obj

        return parent, ref, mode


class ContextBuilder:
    """
    Builds the standardized, context-aware JSON response for tools.
    """

    def __init__(self, session):
        self.session = session
        self.visualizer = ContextVisualizer(session)

    def build_response_data(self, element: Any, element_id: str) -> dict:
        """
        Build the 'data' dictionary for the response.
        Includes cursor info, visual tree, and path.
        """
        # Get parent ID
        parent_id = "?"
        parent = ElementNavigator.get_docx_parent(element, self.session.document)
        if parent:
            if parent == self.session.document:
                parent_id = "document_body"
            else:
                parent_id = self.session._get_element_id(parent, auto_register=True) or "?"

        # Get Path
        path = ElementNavigator.get_path(element)

        # Generate Visual Tree
        # Default to concise mode (range=1)
        visual = self.visualizer.generate_tree_view(element, sibling_range=1)

        # Get position from session if available
        position = "after"
        if hasattr(self.session, 'cursor'):
             position = self.session.cursor.position

        return {
            "element_id": element_id,
            "cursor": {
                "element_id": element_id,
                "parent_id": parent_id,
                "position": position,
                "path": path,
                "visual": visual,
                "context": visual  # Alias visual as context for compatibility
            }
        }

