"""Formatting and styling tools"""
import json
import logging
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches
from docx_editor_skill.core.properties import set_properties
from docx_editor_skill.utils.session_helpers import get_active_session
from docx_editor_skill.core.format_painter import FormatPainter
from docx_editor_skill.utils.format_template import TemplateManager
from docx_editor_skill.core.response import (
    create_markdown_response,
    create_error_response,
    create_context_aware_response,
    create_success_response
)

logger = logging.getLogger(__name__)

# Helper for alignment
ALIGNMENT_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def docx_set_alignment(paragraph_id: str, alignment: str) -> str:
    """
    Set horizontal alignment for a paragraph.

    Controls how text is aligned within the paragraph margins. Affects the entire
    paragraph, not individual runs.

    Typical Use Cases:
        - Center headings and titles
        - Right-align dates or signatures
        - Justify body text for formal documents

    Args:        paragraph_id (str): ID of the paragraph to align.
        alignment (str): Alignment value. Must be one of:
            - "left": Align text to left margin (default)
            - "center": Center text between margins
            - "right": Align text to right margin
            - "justify": Distribute text evenly between margins

    Returns:
        str: JSON response with success message and cursor context.

    Raises:
        ValueError: If session_id or paragraph_id is invalid.
        ValueError: If alignment is not one of the valid values.

    Examples:
        Center a heading:
        >>> session_id = docx_create()
        >>> para_id = docx_insert_heading("Title", position="end:document_body", level=0)
        >>> docx_set_alignment(para_id, "center")

        Right-align a date:
        >>> para_id = docx_insert_paragraph("January 21, 2026", position="end:document_body")
        >>> docx_set_alignment(para_id, "right")

        Justify body text:
        >>> para_id = docx_insert_paragraph("Long paragraph text...", position="end:document_body")
        >>> docx_set_alignment(para_id, "justify")

    Notes:
        - Alignment is case-insensitive
        - Applies to entire paragraph, not individual runs
        - Default alignment is "left"

    See Also:
        - docx_insert_paragraph: Create paragraph to align
        - docx_set_properties: Advanced paragraph formatting
    """


    session, error = get_active_session()
    if error:
        return error
    logger.debug(f"docx_set_alignment called: session_id={session.session_id}, paragraph_id={paragraph_id}, alignment={alignment}")

    try:


        paragraph = session.get_object(paragraph_id)


    except ValueError as e:


        if "Special ID" in str(e) or "not available" in str(e):


            return create_error_response(str(e), error_type="SpecialIDNotAvailable")


        raise

    if not paragraph:
        logger.error(f"docx_set_alignment failed: Paragraph {paragraph_id} not found")
        return create_error_response(f"Paragraph {paragraph_id} not found", error_type="ElementNotFound")

    if alignment.lower() not in ALIGNMENT_MAP:
        logger.error(f"docx_set_alignment failed: Invalid alignment {alignment}")
        return create_error_response(f"Invalid alignment: {alignment}. Must be one of {list(ALIGNMENT_MAP.keys())}", error_type="ValidationError")

    try:
        paragraph.alignment = ALIGNMENT_MAP[alignment.lower()]

        # Update cursor to point to this paragraph
        session.cursor.element_id = paragraph_id
        session.cursor.position = "after"

        logger.debug(f"docx_set_alignment success: {paragraph_id}")
        return create_context_aware_response(
            session,
            message=f"Alignment set to {alignment} for {paragraph_id}",
            element_id=paragraph_id,
            alignment=alignment
        )
    except Exception as e:
        logger.exception(f"docx_set_alignment failed: {e}")
        return create_error_response(f"Failed to set alignment: {str(e)}", error_type="UpdateError")

def docx_set_properties(properties: str, element_id: str = None) -> str:
    """
    Set advanced properties on a document element using JSON configuration.

    Provides flexible property setting for runs, paragraphs, tables, and cells.
    Supports font, paragraph format, and other advanced formatting options.

    Typical Use Cases:
        - Apply complex formatting in one call
        - Set properties not covered by dedicated tools
        - Batch apply multiple formatting options

    Args:        properties (str): JSON string defining properties to set. Structure:
            {
                "font": {"bold": true, "size": 12, "name": "Arial"},
                "paragraph_format": {"alignment": "center", "line_spacing": 1.5}
            }
        element_id (str, optional): ID of element to modify (run, paragraph, table, cell).
            If None, uses last accessed object from session context.

    Returns:
        str: JSON response with success message and cursor context.

    Raises:
        ValueError: If session_id or element_id is invalid.
        ValueError: If properties is not valid JSON.
        ValueError: If no element context available when element_id is None.

    Examples:
        Set font properties:
        >>> session_id = docx_create()
        >>> para_id = docx_insert_paragraph("", position="end:document_body")
        >>> run_id = docx_insert_run("Text", position=f"inside:{para_id}")
        >>> props = '{"font": {"bold": true, "size": 14, "name": "Arial"}}'
        >>> docx_set_properties(props, element_id=run_id)

        Set paragraph format:
        >>> props = '{"paragraph_format": {"alignment": "center", "line_spacing": 1.5}}'
        >>> docx_set_properties(props, element_id=para_id)

    Notes:
        - JSON must be valid and properly escaped
        - Available properties depend on element type
        - Use dedicated tools (docx_set_font, docx_set_alignment) for simpler cases

    See Also:
        - docx_set_font: Simpler font formatting
        - docx_set_alignment: Simpler alignment setting
    """


    session, error = get_active_session()
    if error:
        return error
    logger.debug(f"docx_set_properties called: session_id={session.session_id}, element_id={element_id}, properties_len={len(properties)}")

    target_id = element_id
    if not target_id:
        target_id = session.last_accessed_id
        if not target_id:
            logger.error("docx_set_properties failed: No element context available")
            return create_error_response("No element context available. Please specify element_id.", error_type="NoContext")

    try:


        obj = session.get_object(target_id)


    except ValueError as e:


        if "Special ID" in str(e) or "not available" in str(e):


            return create_error_response(str(e), error_type="SpecialIDNotAvailable")


        raise

    if not obj:
        logger.error(f"docx_set_properties failed: Object {target_id} not found")
        return create_error_response(f"Object {target_id} not found", error_type="ElementNotFound")

    try:
        props_dict = json.loads(properties)
    except json.JSONDecodeError as e:
        logger.error(f"docx_set_properties failed: Invalid JSON - {e}")
        return create_error_response("Invalid JSON string in properties", error_type="JSONDecodeError")

    try:
        set_properties(obj, props_dict)

        # Update context: we accessed/modified this object
        session.update_context(target_id, action="access")

        # Update cursor
        session.cursor.element_id = target_id
        session.cursor.position = "after"

        logger.debug(f"docx_set_properties success: {target_id}")
        return create_context_aware_response(
            session,
            message=f"Properties updated for {target_id}",
            element_id=target_id
        )
    except Exception as e:
        logger.exception(f"docx_set_properties failed: {e}")
        return create_error_response(f"Failed to set properties: {str(e)}", error_type="UpdateError")

def docx_format_copy(source_id: str, target_id: str) -> str:
    """
    Copy formatting from source element to target element.

    Transfers formatting properties between elements of the same type. Supports
    Run-to-Run, Paragraph-to-Paragraph, and Table-to-Table copying.

    Typical Use Cases:
        - Apply consistent formatting across elements
        - Replicate styling from templates
        - Maintain formatting consistency

    Args:        source_id (str): ID of source element to copy formatting from.
        target_id (str): ID of target element to apply formatting to.

    Returns:
        str: JSON response with success message and cursor context.

    Raises:
        ValueError: If session_id, source_id, or target_id is invalid.
        ValueError: If source and target are not the same element type.

    Examples:
        Copy run formatting:
        >>> session_id = docx_create()
        >>> para_id = docx_insert_paragraph("", position="end:document_body")
        >>> run1 = docx_insert_run("Styled", position=f"inside:{para_id}")
        >>> docx_set_font(run1, bold=True, size=14, color_hex="FF0000")
        >>> run2 = docx_insert_run("Copy style", position=f"inside:{para_id}")
        >>> docx_format_copy(run1, run2)

        Copy paragraph formatting:
        >>> para1 = docx_insert_paragraph("Source", position="end:document_body")
        >>> docx_set_alignment(para1, "center")
        >>> para2 = docx_insert_paragraph("Target", position="end:document_body")
        >>> docx_format_copy(para1, para2)

    Notes:
        - Source and target must be same type (run/paragraph/table)
        - Only formatting is copied, not content
        - Target content is preserved
        - Supports: font properties, alignment, spacing, borders

    See Also:
        - docx_set_font: Set specific font properties
        - docx_set_alignment: Set alignment directly
    """


    session, error = get_active_session()
    if error:
        return error
    logger.debug(f"docx_format_copy called: session_id={session.session_id}, source_id={source_id}, target_id={target_id}")

    try:


        source = session.get_object(source_id)


    except ValueError as e:


        if "Special ID" in str(e) or "not available" in str(e):


            return create_error_response(str(e), error_type="SpecialIDNotAvailable")


        raise

    if not source:
        logger.error(f"docx_format_copy failed: Source object {source_id} not found")
        return create_error_response(f"Source object {source_id} not found", error_type="ElementNotFound")

    try:


        target = session.get_object(target_id)


    except ValueError as e:


        if "Special ID" in str(e) or "not available" in str(e):


            return create_error_response(str(e), error_type="SpecialIDNotAvailable")


        raise

    if not target:
        logger.error(f"docx_format_copy failed: Target object {target_id} not found")
        return create_error_response(f"Target object {target_id} not found", error_type="ElementNotFound")

    painter = FormatPainter()
    try:
        painter.copy_format(source, target)

        # Update context: we modified the target
        session.update_context(target_id, action="access")

        # Update cursor
        session.cursor.element_id = target_id
        session.cursor.position = "after"

        logger.debug(f"docx_format_copy success: {source_id} -> {target_id}")
        return create_context_aware_response(
            session,
            message=f"Format copied from {source_id} to {target_id}",
            element_id=target_id,
            source_id=source_id
        )
    except Exception as e:
        logger.exception(f"docx_format_copy failed: {e}")
        return create_error_response(f"Failed to copy format: {str(e)}", error_type="CopyError")

def docx_set_margins(
    top: float = None,
    bottom: float = None,
    left: float = None,
    right: float = None
) -> str:
    """
    Set page margins for the document section.

    Adjusts the white space around the content area on each page. Applies to the
    last section of the document (typically the entire document if single-section).

    Typical Use Cases:
        - Adjust margins for printing requirements
        - Create narrow margins for more content
        - Set wide margins for annotations

    Args:        top (float, optional): Top margin in inches.
        bottom (float, optional): Bottom margin in inches.
        left (float, optional): Left margin in inches.
        right (float, optional): Right margin in inches.

    Returns:
        str: JSON response with success message.

    Raises:
        ValueError: If session_id is invalid or session has expired.

    Examples:
        Set all margins:
        >>> session_id = docx_create()
        >>> docx_set_margins(top=1.0, bottom=1.0, left=1.0, right=1.0)

        Set narrow margins:
        >>> docx_set_margins(top=0.5, bottom=0.5, left=0.5, right=0.5)

        Set only top and bottom:
        >>> docx_set_margins(top=1.5, bottom=1.5)

    Notes:
        - All measurements are in inches (1 inch = 2.54 cm)
        - Omitted parameters leave existing margins unchanged
        - Applies to the last section (usually entire document)
        - Default Word margins are typically 1 inch on all sides

    See Also:
        - docx_insert_page_break: Control page breaks
    """


    session, error = get_active_session()
    if error:
        return error
    logger.debug(f"docx_set_margins called: session_id={session.session_id}, top={top}, bottom={bottom}, left={left}, right={right}")

    try:
        # For simplicity, apply to the last section (usually where we are working)
        section = session.document.sections[-1]

        if top is not None:
            section.top_margin = Inches(top)
        if bottom is not None:
            section.bottom_margin = Inches(bottom)
        if left is not None:
            section.left_margin = Inches(left)
        if right is not None:
            section.right_margin = Inches(right)

        logger.debug("docx_set_margins success")
        return create_markdown_response(
            session=session,
            message="Margins updated successfully",
            operation="Operation",
            show_context=True,
            include_cursor=False,
            margins=json.dumps({"top": top, "bottom": bottom, "left": left, "right": right})
        )
    except Exception as e:
        logger.exception(f"docx_set_margins failed: {e}")
        return create_error_response(f"Failed to set margins: {str(e)}", error_type="UpdateError")

def docx_extract_format_template(element_id: str) -> str:
    """
    Extract style and formatting properties from an element as a reusable template.

    Captures font, paragraph, or table properties into a JSON structure that can be
    stored and applied to other elements later.

    Typical Use Cases:
        - Create a library of reusable styles
        - "Pickup" formatting to apply elsewhere
        - Persist styles between sessions

    Args:        element_id (str): ID of the element to extract format from.

    Returns:
        str: JSON response containing the format template.

    Raises:
        ValueError: If element not found or type unsupported.
    """


    session, error = get_active_session()
    if error:
        return error
    logger.debug(f"docx_extract_format_template called: session_id={session.session_id}, element_id={element_id}")

    try:


        element = session.get_object(element_id)


    except ValueError as e:


        if "Special ID" in str(e) or "not available" in str(e):


            return create_error_response(str(e), error_type="SpecialIDNotAvailable")


        raise

    if not element:
        logger.error(f"docx_extract_format_template failed: Element {element_id} not found")
        return create_error_response(f"Element {element_id} not found", error_type="ElementNotFound")

    manager = TemplateManager()
    try:
        template = manager.extract_template(element)
        template_str = manager.to_json(template)

        logger.debug(f"docx_extract_format_template success: {element_id}")
        # We return the raw JSON string as the main result for compatibility?
        # No, we promised structured JSON.
        # But wait, the return type annotation is `str`.
        # If we return a JSON string that contains the template JSON string inside it, or as an object?
        # Ideally as an object in 'data'.
        return create_success_response(
            message=f"Template extracted from {element_id}",
            template=json.loads(template_str),
            source_id=element_id
        )
    except Exception as e:
        logger.exception(f"docx_extract_format_template failed: {e}")
        return create_error_response(f"Failed to extract template: {str(e)}", error_type="ExtractionError")


def docx_apply_format_template(element_id: str, template_json: str) -> str:
    """
    Apply a format template to an element.

    Applies properties defined in the template JSON to the target element.
    Ignores properties that don't apply to the target type.

    Typical Use Cases:
        - Apply standard styles
        - Replicate formatting extracted earlier
        - Batch format elements

    Args:        element_id (str): Target element ID.
        template_json (str): JSON string returned by docx_extract_format_template.

    Returns:
        str: JSON response with success message.
    """


    session, error = get_active_session()
    if error:
        return error
    logger.debug(f"docx_apply_format_template called: session_id={session.session_id}, element_id={element_id}, template_len={len(template_json)}")

    try:


        element = session.get_object(element_id)


    except ValueError as e:


        if "Special ID" in str(e) or "not available" in str(e):


            return create_error_response(str(e), error_type="SpecialIDNotAvailable")


        raise

    if not element:
        logger.error(f"docx_apply_format_template failed: Element {element_id} not found")
        return create_error_response(f"Element {element_id} not found", error_type="ElementNotFound")

    manager = TemplateManager()
    try:
        template = manager.from_json(template_json)
        manager.apply_template(element, template)

        # Update context
        session.update_context(element_id, action="access")

        # Update cursor
        session.cursor.element_id = element_id
        session.cursor.position = "after"

        logger.debug(f"docx_apply_format_template success: {element_id}")
        return create_context_aware_response(
            session,
            message=f"Template applied to {element_id}",
            element_id=element_id
        )
    except Exception as e:
        logger.exception(f"docx_apply_format_template failed: {e}")
        return create_error_response(f"Failed to apply template: {str(e)}", error_type="ApplicationError")