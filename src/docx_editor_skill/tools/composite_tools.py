"""Composite tools for common scenarios - high-level operations"""
import logging
import json
import re
from typing import Optional

from docx_editor_skill.core.response import create_markdown_response, create_error_response
from docx_editor_skill.utils.session_helpers import get_active_session

logger = logging.getLogger(__name__)


def _extract_element_id(response: str) -> str:
    """
    Extract element_id from Markdown or JSON response.

    Args:
        response: Markdown or JSON response string

    Returns:
        str: Element ID
    """
    # Try Markdown format first (new format)
    match = re.search(r'\*\*Element ID\*\*:\s*(\w+)', response)
    if match:
        return match.group(1)

    # Fallback to JSON format (legacy)
    try:
        data = json.loads(response)
        if isinstance(data, dict) and "data" in data and "element_id" in data["data"]:
            return data["data"]["element_id"]
        return response
    except (json.JSONDecodeError, KeyError):
        return response


def _extract_metadata_field(response: str, field_name: str) -> Optional[str]:
    """
    Extract a metadata field from Markdown response.

    Args:
        response: Markdown response string
        field_name: Field name to extract (e.g., "rows_filled")

    Returns:
        str: Field value or None if not found
    """
    display_name = field_name.replace('_', ' ').title()
    pattern = rf'\*\*{re.escape(display_name)}\*\*:\s*(.+?)(?:\n|$)'
    match = re.search(pattern, response)
    if match:
        return match.group(1).strip()
    return None


def docx_insert_formatted_paragraph(
    text: str,
    position: str,
    bold: bool = False,
    italic: bool = False,
    size: float = None,
    color_hex: str = None,
    alignment: str = None,
    style: str = None
) -> str:
    """
    Create a formatted paragraph in one step.

    Combines docx_insert_paragraph + docx_insert_run + docx_set_font + docx_set_alignment
    into a single operation for common use cases.

    Typical Use Cases:
        - Quickly add formatted content to documents
        - Create styled headings or emphasis text
        - Reduce multiple tool calls to one

    Args:        text (str): Paragraph text content.
        position (str): Insertion position string (e.g., "after:para_123").
        bold (bool): Make text bold. Defaults to False.
        italic (bool): Make text italic. Defaults to False.
        size (float): Font size in points (e.g., 12.0, 14.5).
        color_hex (str): Hex color without '#' (e.g., "FF0000" for red).
        alignment (str): Paragraph alignment: "left", "center", "right", "justify".
        style (str): Built-in style name (e.g., 'Body Text').

    Returns:
        str: Element ID of the created paragraph.

    Examples:
        Create bold red text:
        >>> para_id = docx_insert_formatted_paragraph(
        ...     session_id, "Important!", position="end:document_body", bold=True, color_hex="FF0000"
        ... )

        Create centered heading:
        >>> para_id = docx_insert_formatted_paragraph(
        ...     session_id, "Title", position="end:document_body", size=16, alignment="center"
        ... )
    """
    from docx_editor_skill.tools.paragraph_tools import docx_insert_paragraph
    from docx_editor_skill.tools.run_tools import docx_insert_run, docx_set_font
    from docx_editor_skill.tools.format_tools import docx_set_alignment

    try:
        # Step 1: Create paragraph
        para_response = docx_insert_paragraph("", position=position, style=style)
        para_id = _extract_element_id(para_response)

        # Step 2: Add run with text
        run_response = docx_insert_run(text, position=f"inside:{para_id}")
        run_id = _extract_element_id(run_response)

        # Step 3: Apply font formatting if specified
        if any([bold, italic, size, color_hex]):
            docx_set_font(run_id, size=size, bold=bold, italic=italic, color_hex=color_hex)

        # Step 4: Apply alignment if specified
        if alignment:
            docx_set_alignment(para_id, alignment)

        logger.info(f"Created formatted paragraph: {para_id}")
        return para_id

    except Exception as e:
        logger.exception(f"Failed to create formatted paragraph: {e}")
        raise ValueError(f"Failed to create formatted paragraph: {e}")


def docx_quick_edit(
    search_text: str,
    new_text: str = None,
    bold: bool = None,
    italic: bool = None,
    size: float = None,
    color_hex: str = None
) -> str:
    """
    Find and edit paragraphs in one step.

    Combines docx_find_paragraphs + docx_update_paragraph_text + formatting
    for quick editing workflows.

    Typical Use Cases:
        - Edit existing documents quickly
        - Find and replace with formatting changes
        - Update specific content without manual ID tracking

    Args:        search_text (str): Text to search for in paragraphs.
        new_text (str): New text to replace (if None, only formatting changes).
        bold (bool): Set bold formatting.
        italic (bool): Set italic formatting.
        size (float): Font size in points.
        color_hex (str): Hex color without '#'.

    Returns:
        str: JSON with modified paragraph count and IDs.

    Examples:
        Replace text:
        >>> result = docx_quick_edit("old text", new_text="new text")

        Change formatting only:
        >>> result = docx_quick_edit("important", bold=True, color_hex="FF0000")
    """
    from docx_editor_skill.tools.content_tools import docx_find_paragraphs
    from docx_editor_skill.tools.paragraph_tools import docx_update_paragraph_text
    from docx_editor_skill.tools.run_tools import docx_set_font

    try:
        # Find matching paragraphs
        matches_md = docx_find_paragraphs(search_text)

        # Parse Markdown response
        if "No matching paragraphs found" in matches_md:
            # Return Markdown format
            md_lines = ["# Quick Edit Result\n"]
            md_lines.append("**Modified Count**: 0")
            md_lines.append("\n**Modified Paragraph IDs**: None")
            return "\n".join(md_lines)

        # Extract matches from Markdown
        matches = []
        match_pattern = r'## Match \d+\s*\n\*\*ID\*\*:\s*`([^`]+)`\s*\n\*\*Text\*\*:\s*(.+?)(?=\n##|\n\*\*Context|\Z)'
        for match in re.finditer(match_pattern, matches_md, re.DOTALL):
            para_id = match.group(1).strip()
            text = match.group(2).strip()
            matches.append({'id': para_id, 'text': text})

        if not matches:
            # Return Markdown format
            md_lines = ["# Quick Edit Result\n"]
            md_lines.append("**Modified Count**: 0")
            md_lines.append("\n**Modified Paragraph IDs**: None")
            return "\n".join(md_lines)

        modified_ids = []

        for match in matches:
            para_id = match["id"]

            # Update text if specified
            if new_text is not None:
                docx_update_paragraph_text(para_id, new_text)

            # Apply formatting if specified
            if any([bold is not None, italic is not None, size, color_hex]):
                session, error = get_active_session()
                paragraph = session.get_object(para_id)

                # Apply formatting to all runs in paragraph
                for run in paragraph.runs:
                    run_id = session._get_element_id(run, auto_register=True)
                    docx_set_font(run_id, size=size, bold=bold, italic=italic, color_hex=color_hex)

            modified_ids.append(para_id)

        logger.info(f"Quick edit modified {len(modified_ids)} paragraphs")

        # Return Markdown format
        md_lines = ["# Quick Edit Result\n"]
        md_lines.append(f"**Modified Count**: {len(modified_ids)}")
        md_lines.append("\n**Modified Paragraph IDs**:")
        for pid in modified_ids:
            md_lines.append(f"- `{pid}`")

        return "\n".join(md_lines)

    except Exception as e:
        logger.exception(f"Quick edit failed: {e}")
        raise ValueError(f"Quick edit failed: {e}")


def docx_get_structure_summary(
    max_headings: int = 10,
    max_tables: int = 5,
    max_paragraphs: int = 0,
    include_content: bool = False
) -> str:
    """
    Get a lightweight document structure summary.

    Returns a simplified structure with configurable limits to reduce token usage.
    Much lighter than docx_extract_template_structure for quick overview.

    Typical Use Cases:
        - Quickly understand document layout
        - Get table of contents
        - Find specific sections without full content

    Args:        max_headings (int): Maximum headings to return. Defaults to 10.
        max_tables (int): Maximum tables to return. Defaults to 5.
        max_paragraphs (int): Maximum paragraphs to return. Defaults to 0 (none).
        include_content (bool): Include text content. Defaults to False (structure only).

    Returns:
        str: JSON with document structure summary.

    Examples:
        Get headings only:
        >>> summary = docx_get_structure_summary()

        Get headings and tables:
        >>> summary = docx_get_structure_summary(max_tables=10)
    """

    session, error = get_active_session()
    if error:
        return error

    try:
        doc = session.document
        structure = {
            "headings": [],
            "tables": [],
            "paragraphs": []
        }

        heading_count = 0
        table_count = 0
        para_count = 0

        for element in doc.element.body:
            # Process headings
            if element.tag.endswith('p'):
                para = next((p for p in doc.paragraphs if p._element == element), None)
                if para and para.style.name.startswith('Heading'):
                    if heading_count < max_headings:
                        heading_info = {
                            "level": int(para.style.name.split()[-1]) if para.style.name.split()[-1].isdigit() else 1,
                            "style": para.style.name,
                            "element_id": session._get_element_id(para, auto_register=True)
                        }
                        if include_content:
                            heading_info["text"] = para.text
                        structure["headings"].append(heading_info)
                        heading_count += 1

                # Process regular paragraphs
                elif para and max_paragraphs > 0 and para_count < max_paragraphs:
                    para_info = {
                        "style": para.style.name,
                        "element_id": session._get_element_id(para, auto_register=True)
                    }
                    if include_content:
                        para_info["text"] = para.text[:100]  # Truncate long text
                    structure["paragraphs"].append(para_info)
                    para_count += 1

            # Process tables
            elif element.tag.endswith('tbl') and table_count < max_tables:
                table = next((t for t in doc.tables if t._element == element), None)
                if table:
                    table_info = {
                        "rows": len(table.rows),
                        "cols": len(table.columns),
                        "element_id": session._get_element_id(table, auto_register=True)
                    }
                    if include_content:
                        # Include first row as header sample
                        table_info["first_row"] = [cell.text[:50] for cell in table.rows[0].cells]
                    structure["tables"].append(table_info)
                    table_count += 1

        structure["summary"] = {
            "total_headings": heading_count,
            "total_tables": table_count,
            "total_paragraphs": para_count
        }

        logger.info(f"Generated structure summary for session {session.session_id}")

        # Return Markdown format
        md_lines = ["# Document Structure Summary\n"]
        md_lines.append("## Summary")
        md_lines.append(f"- **Total Headings**: {heading_count}")
        md_lines.append(f"- **Total Tables**: {table_count}")
        md_lines.append(f"- **Total Paragraphs**: {para_count}")
        md_lines.append("")

        if structure["headings"]:
            md_lines.append("## Headings\n")
            for h in structure["headings"]:
                level = h.get("level", 1)
                text = h.get("text", "") if include_content else "[content hidden]"
                md_lines.append(f"{'#' * (level + 1)} {text}")
                md_lines.append(f"*ID: `{h['element_id']}`*\n")

        if structure["tables"]:
            md_lines.append("## Tables\n")
            for idx, t in enumerate(structure["tables"], 1):
                md_lines.append(f"### Table {idx}")
                md_lines.append(f"- **Size**: {t['rows']}x{t['cols']}")
                md_lines.append(f"- **ID**: `{t['element_id']}`")
                if include_content and t.get("first_row"):
                    md_lines.append(f"- **First Row**: {', '.join(t['first_row'])}")
                md_lines.append("")

        if structure["paragraphs"]:
            md_lines.append("## Paragraphs\n")
            for idx, p in enumerate(structure["paragraphs"], 1):
                md_lines.append(f"### Paragraph {idx}")
                md_lines.append(f"- **Style**: {p['style']}")
                md_lines.append(f"- **ID**: `{p['element_id']}`")
                if include_content and p.get("text"):
                    md_lines.append(f"- **Text**: {p['text']}")
                md_lines.append("")

        return "\n".join(md_lines)

    except Exception as e:
        logger.exception(f"Failed to generate structure summary: {e}")
        raise ValueError(f"Failed to generate structure summary: {e}")


def docx_smart_fill_table(
    table_identifier: str,
    data: str,
    has_header: bool = True,
    auto_resize: bool = True,
    preserve_formatting: bool = False
) -> str:
    """
    Intelligently fill table with data, auto-expanding rows as needed.

    Combines table lookup + fill + row addition in one operation.

    Typical Use Cases:
        - Fill tables from database query results
        - Import CSV/JSON data into documents
        - Populate templates with dynamic data

    Args:        table_identifier (str): Table index (e.g., "0") or text to find table.
        data (str): JSON array of arrays: [["col1", "col2"], ["val1", "val2"]].
        has_header (bool): First row is header. Defaults to True.
        auto_resize (bool): Auto-add rows if needed. Defaults to True.

    Returns:
        str: JSON with fill status and row count.

    Examples:
        Fill table by index:
        >>> data = '[["Name", "Age"], ["Alice", "30"], ["Bob", "25"]]'
        >>> result = docx_smart_fill_table("0", data)

        Fill table by content:
        >>> result = docx_smart_fill_table("Employee", data)
    """
    from docx_editor_skill.tools.table_tools import (
        docx_get_table, docx_find_table, docx_fill_table, docx_insert_table_row
    )

    session, error = get_active_session()
    if error:
        return error

    try:
        # Parse data
        data_array = json.loads(data)
        if not isinstance(data_array, list) or not data_array:
            raise ValueError("Data must be a non-empty JSON array")

        # Find table - try multiple strategies
        table_id = None
        table = None

        # Strategy 1: Check if it's already a valid table ID
        if table_identifier.startswith("table_"):
            table = session.get_object(table_identifier)
            if table is not None:
                table_id = table_identifier
                logger.info(f"Using existing table ID: {table_id}")

        # Strategy 2: Try as index
        if table is None:
            try:
                table_response = docx_get_table(int(table_identifier))
                table_id = _extract_element_id(table_response)
                table = session.get_object(table_id)
                logger.info(f"Found table by index: {table_id}")
            except (ValueError, TypeError):
                pass

        # Strategy 3: Try as search text
        if table is None:
            table_response = docx_find_table(table_identifier)
            table_id = _extract_element_id(table_response)
            table = session.get_object(table_id)
            logger.info(f"Found table by search text: {table_id}")

        if table is None:
            raise ValueError(f"Table not found with identifier: {table_identifier}")

        # Calculate rows needed
        data_rows = len(data_array) - (1 if has_header else 0)
        existing_rows = len(table.rows) - (1 if has_header else 0)

        # Add rows if needed
        if auto_resize and data_rows > existing_rows:
            rows_to_add = data_rows - existing_rows
            for _ in range(rows_to_add):
                docx_insert_table_row(position=f"inside:{table_id}")
            logger.info(f"Added {rows_to_add} rows to table")

        # Fill table
        start_row = 1 if has_header else 0
        fill_result = docx_fill_table(data, table_id, start_row=start_row, preserve_formatting=preserve_formatting)

        # Extract metadata from Markdown response
        rows_filled_str = _extract_metadata_field(fill_result, "rows_filled")
        rows_filled = int(rows_filled_str) if rows_filled_str else len(data_array)

        logger.info(f"Smart fill completed: rows_filled={rows_filled}, rows_added={max(0, data_rows - existing_rows) if auto_resize else 0}")

        return create_markdown_response(
            session=session,
            message=f"Smart filled table with {rows_filled} rows",
            rows_filled=rows_filled,
            rows_added=max(0, data_rows - existing_rows) if auto_resize else 0,
            preserve_formatting=preserve_formatting
        )

    except Exception as e:
        logger.exception(f"Smart fill table failed: {e}")
        return create_error_response(f"Smart fill table failed: {e}", error_type="SmartFillError")


def docx_format_range(
    start_text: str,
    end_text: str,
    bold: bool = None,
    italic: bool = None,
    size: float = None,
    color_hex: str = None
) -> str:
    """
    Apply formatting to a range of paragraphs between two markers.

    Finds paragraphs containing start_text and end_text, then formats
    all paragraphs in between (inclusive).

    Typical Use Cases:
        - Format entire sections at once
        - Apply consistent styling to ranges
        - Bulk formatting operations

    Args:        start_text (str): Text marking range start.
        end_text (str): Text marking range end.
        bold (bool): Set bold formatting.
        italic (bool): Set italic formatting.
        size (float): Font size in points.
        color_hex (str): Hex color without '#'.

    Returns:
        str: JSON with formatted paragraph count.

    Examples:
        Format section:
        >>> result = docx_format_range(
        ...     session_id, "Chapter 1", "Chapter 2", bold=True, size=14
        ... )
    """
    from docx_editor_skill.tools.run_tools import docx_set_font

    session, error = get_active_session()
    if error:
        return error

    try:
        doc = session.document
        paragraphs = list(doc.paragraphs)

        # Find start and end indices
        start_idx = None
        end_idx = None

        for i, para in enumerate(paragraphs):
            if start_text in para.text and start_idx is None:
                start_idx = i
            if end_text in para.text:
                end_idx = i

        if start_idx is None:
            raise ValueError(f"Start text '{start_text}' not found")
        if end_idx is None:
            raise ValueError(f"End text '{end_text}' not found")
        if start_idx > end_idx:
            raise ValueError("Start text appears after end text")

        # Format range
        formatted_count = 0
        for i in range(start_idx, end_idx + 1):
            para = paragraphs[i]
            for run in para.runs:
                run_id = session._get_element_id(run, auto_register=True)
                docx_set_font(run_id, size=size, bold=bold, italic=italic, color_hex=color_hex)
            formatted_count += 1

        logger.info(f"Formatted range: {formatted_count} paragraphs")

        # Return Markdown format
        md_lines = ["# Format Range Result\n"]
        md_lines.append(f"**Formatted Count**: {formatted_count}")
        md_lines.append(f"**Start Index**: {start_idx}")
        md_lines.append(f"**End Index**: {end_idx}")

        return "\n".join(md_lines)

    except Exception as e:
        logger.exception(f"Format range failed: {e}")
        raise ValueError(f"Format range failed: {e}")
