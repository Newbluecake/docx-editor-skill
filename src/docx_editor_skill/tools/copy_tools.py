"""Copy and Metadata tools"""
import logging

from docx_editor_skill.utils.copy_engine import CopyEngine
from docx_editor_skill.utils.metadata_tools import MetadataTools
from docx_editor_skill.services.navigation import PositionResolver
from docx_editor_skill.core.xml_util import ElementManipulator
from docx_editor_skill.utils.session_helpers import get_active_session

logger = logging.getLogger(__name__)


def docx_get_element_source(element_id: str) -> str:
    """
    Get the source lineage metadata of an element.

    Returns metadata about where this element was copied from, if applicable.
    Useful for tracking content origin in document generation workflows.

    Args:
        session_id (str): Active session ID.
        element_id (str): ID of the element to check.

    Returns:
        str: JSON string containing source metadata, or empty object if none.
    """

    session, error = get_active_session()
    if error:
        return error

    logger.debug(f"docx_get_element_source called: session_id={session.session_id}, element_id={element_id}")

    metadata = session.get_metadata(element_id)
    if not metadata:
        logger.debug(f"docx_get_element_source success: no metadata for {element_id}")
        return "No source metadata found for this element."

    logger.debug(f"docx_get_element_source success: found metadata for {element_id}")

    # Return Markdown format
    md_lines = ["# Element Source Metadata\n"]
    md_lines.append(f"**Element ID**: `{element_id}`\n")

    for key, value in metadata.items():
        key_display = key.replace('_', ' ').title()
        md_lines.append(f"**{key_display}**: {value}")

    return "\n".join(md_lines)


def docx_copy_elements_range(start_id: str, end_id: str, position: str) -> str:
    """
    Copy a range of elements (e.g., from one heading to another) to a target location.

    Copies all supported elements (paragraphs, tables) between the start and end elements
    (inclusive). Maintains the relative order and structure.

    Typical Use Cases:
        - Copy entire chapters or sections
        - Duplicate complex document structures
        - Merge content from different parts of a document

    Args:
        session_id (str): Active session ID.
        start_id (str): ID of the first element in the range.
        end_id (str): ID of the last element in the range.
        position (str): Insertion position string (e.g., "after:para_123").

    Returns:
        str: JSON array of new element IDs mapped to their original source IDs.
            Example: [{"source": "para_1", "new": "para_99"}, ...]

    Raises:
        ValueError: If elements are invalid or not siblings.
    """

    session, error = get_active_session()
    if error:
        return error

    logger.debug(f"docx_copy_elements_range called: session_id={session.session_id}, start_id={start_id}, end_id={end_id}, position={position}")

    try:
        start_el = session.get_object(start_id)
    except ValueError as e:
        if "Special ID" in str(e) or "not available" in str(e):
            raise ValueError(f"Start element: {str(e)}")
        raise

    try:
        end_el = session.get_object(end_id)
    except ValueError as e:
        if "Special ID" in str(e) or "not available" in str(e):
            raise ValueError(f"End element: {str(e)}")
        raise

    if not start_el or not end_el:
        logger.error("docx_copy_elements_range failed: Start or end element not found")
        raise ValueError("Start or end element not found")

    try:
        resolver = PositionResolver(session)
        target_parent, ref_element, mode = resolver.resolve(position, default_parent=session.document)
    except ValueError as e:
        raise ValueError(str(e))

    engine = CopyEngine()
    try:
        # Perform range copy
        # We append to target_parent
        ref_xml = None
        if mode == "after" and ref_element:
            ref_xml = ref_element._element
        elif mode == "before" and ref_element:
            ref_xml = ref_element._element.getprevious()

        new_objects = engine.copy_range(start_el, end_el, target_parent, ref_element_xml=ref_xml)

        if mode in ["start", "before"] and ref_xml is None:
            container_xml = target_parent._element
            if hasattr(target_parent, '_body'):
                container_xml = target_parent._body._element
            for obj in reversed(new_objects):
                obj._element.getparent().remove(obj._element)
                ElementManipulator.insert_at_index(container_xml, obj._element, 0)

        # Register all new objects
        result_map = []
        # We need to map back to source?
        # copy_range returns a list of new objects.
        # We know they correspond 1:1 to the range logic, but we need to re-fetch the source range
        # to map IDs correctly if we want to return a mapping.
        # For MVP, let's just register them and return the list of new IDs.
        # Or better: track source IDs.

        # Re-fetch source elements to map IDs (assuming stability)
        source_elements = engine.get_elements_between(start_el, end_el)

        if len(source_elements) != len(new_objects):
            logger.warning("Source and new object count mismatch during registration")

        for i, new_obj in enumerate(new_objects):
            # Attempt to determine type for prefix
            prefix = "obj"
            if hasattr(new_obj, "add_run"):
                prefix = "para"
            elif hasattr(new_obj, "rows"):
                prefix = "table"

            # Metadata using shared utility
            meta = MetadataTools.create_copy_metadata(operation_type="range_copy")

            # If we can map to source
            if i < len(source_elements):
                # We can't easily get the ID of the source_element object unless we do a reverse lookup
                # in session.object_registry. This is expensive (O(N)).
                # For now, we skip precise source ID mapping in the return value
                # unless we want to implement reverse lookup.
                pass

            new_id = session.register_object(new_obj, prefix, metadata=meta)
            result_map.append({"new_id": new_id, "type": prefix})

        logger.debug(f"docx_copy_elements_range success: copied {len(new_objects)} elements")

        # Return Markdown format
        md_lines = ["# Copy Elements Range Result\n"]
        md_lines.append(f"**Copied Count**: {len(new_objects)}\n")
        md_lines.append("## New Element IDs\n")
        for item in result_map:
            md_lines.append(f"- **{item['type']}**: `{item['new_id']}`")

        return "\n".join(md_lines)

    except Exception as e:
        logger.error(f"docx_copy_elements_range failed: {e}")
        raise ValueError(f"Range copy failed: {str(e)}")
