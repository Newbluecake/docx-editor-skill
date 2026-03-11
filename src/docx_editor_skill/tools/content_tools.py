"""Content reading and search tools"""
import json
import logging
from typing import Optional, Dict, List, Any
from docx_editor_skill.utils.session_helpers import get_active_session
from docx_editor_skill.core.template_parser import TemplateParser

logger = logging.getLogger(__name__)


def docx_read_content(
    max_paragraphs: Optional[int] = None,
    start_from: int = 0,
    include_tables: bool = False,
    return_json: bool = False,
    include_ids: bool = True,
    start_element_id: Optional[str] = None,
    max_tables: Optional[int] = None,
    table_mode: str = "text",
) -> str:
    """
    Read and extract text content from the document with pagination support.

    Extracts text from paragraphs in the document body, preserving order
    but not formatting. Supports limiting output to reduce token usage.

    Typical Use Cases:
        - Preview document content before modification
        - Extract text for analysis or indexing
        - Verify document content after generation
        - Read large documents in chunks

    Args:        max_paragraphs (int, optional): Maximum paragraphs to return. None = all.
        start_from (int, optional): Start from paragraph N (0-based). Defaults to 0.
        include_tables (bool, optional): Include table content. Defaults to False.

    Returns:
        str: Newline-separated text content of paragraphs.
            Returns "[Empty Document]" if document has no content.

    Raises:
        ValueError: If session_id is invalid or session has expired.

    Examples:
        Read all content:
        >>> content = docx_read_content()

        Read first 10 paragraphs:
        >>> content = docx_read_content(max_paragraphs=10)

        Read paragraphs 10-20:
        >>> content = docx_read_content(max_paragraphs=10, start_from=10)

    Notes:
        - Only extracts text, formatting information is not included
        - Empty paragraphs are skipped
        - Use max_paragraphs to limit token usage on large documents

    See Also:
        - docx_find_paragraphs: Search for specific text in paragraphs
        - docx_get_structure_summary: Get lightweight structure overview
    """
    session, error = get_active_session()
    if error:
        return error

    logger.debug(
        f"docx_read_content called: session_id={session.session_id}, max={max_paragraphs}, start={start_from}, "
        f"include_tables={include_tables}, start_element_id={start_element_id}, max_tables={max_tables}, table_mode={table_mode}"
    )

    from docx.table import Table
    from docx.text.paragraph import Paragraph

    # Resolve anchor element (optional)
    anchor_element = None
    if start_element_id:
        obj = session.get_object(start_element_id)
        if obj is not None:
            if isinstance(obj, Paragraph) or isinstance(obj, Table):
                anchor_element = obj._element
            elif hasattr(obj, "_tc"):
                # For cells: climb to parent table
                el = obj._tc
                while el is not None and getattr(el, "tag", "").split("}")[-1] != "tbl":
                    el = el.getparent()
                if el is not None:
                    anchor_element = el

    # Iterate document blocks (paragraphs and tables) in order
    def iter_blocks(doc):
        body = doc.element.body
        for child in body.iterchildren():
            tag = child.tag.split('}')[-1]
            if tag == 'p':
                yield Paragraph(child, doc)
            elif tag == 'tbl':
                yield Table(child, doc)

    blocks = list(iter_blocks(session.document))

    start_index = 0
    if anchor_element is not None:
        for idx, blk in enumerate(blocks):
            if getattr(blk, "_element", None) is anchor_element:
                start_index = idx + 1
                break

    entries: List[Dict[str, Any]] = []
    para_count = 0
    table_count = 0

    for blk in blocks[start_index:]:
        if isinstance(blk, Paragraph):
            if not blk.text.strip():
                continue
            # Only enforce paragraph limit during scan when anchored; otherwise slice later
            if start_element_id is not None and max_paragraphs is not None and para_count >= max_paragraphs:
                continue
            entry: Dict[str, Any] = {"text": blk.text, "type": "paragraph"}
            if include_ids:
                entry["id"] = session._get_element_id(blk, auto_register=True)
            entries.append(entry)
            para_count += 1
        elif include_tables and isinstance(blk, Table):
            if start_element_id is not None and max_tables is not None and table_count >= max_tables:
                continue
            table_id = session._get_element_id(blk, auto_register=True) if include_ids else None
            mode = table_mode if table_mode in ["text", "cells"] else "text"
            if mode != table_mode:
                logger.warning(f"Unsupported table_mode '{table_mode}', fallback to 'text'")

            if mode == "cells":
                table_data = [[cell.text for cell in row.cells] for row in blk.rows]
                text_repr = "\n".join(["\t".join(r) for r in table_data])
            else:
                table_data = None
                text_repr = "\n".join([cell.text for row in blk.rows for cell in row.cells])

            entry: Dict[str, Any] = {"text": text_repr, "type": "table"}
            if table_id is not None:
                entry["id"] = table_id
            if table_data is not None:
                entry["cells"] = table_data
            entry["rows"] = int(len(blk.rows))
            entry["cols"] = int(len(blk.columns)) if blk.rows else 0
            entries.append(entry)
            table_count += 1

    # Apply pagination by index if start_from provided and no anchor
    if start_element_id is None and start_from > 0:
        entries = entries[start_from:]
    if max_paragraphs is not None and start_element_id is None:
        # keep compatibility: limit total entries when no anchor; para limit already applied otherwise
        entries = entries[:max_paragraphs]

    logger.debug(f"docx_read_content success: extracted {len(entries)} entries (paras={para_count}, tables={table_count})")

    if return_json:
        return json.dumps({
            "status": "success",
            "count": len(entries),
            "data": entries
        }, ensure_ascii=False)

    # Return raw Markdown content (not JSON-wrapped)
    if not entries:
        return "[Empty Document]"

    # Build Markdown output
    md_lines = []
    for entry in entries:
        entry_type = entry.get("type", "paragraph")
        text = entry.get("text", "")

        if entry_type == "table":
            # Format table as Markdown table
            rows = entry.get("rows", 0)
            cols = entry.get("cols", 0)
            cells = entry.get("cells")

            if cells:
                # Render as proper Markdown table
                for row_idx, row_data in enumerate(cells):
                    md_lines.append("| " + " | ".join(str(cell) for cell in row_data) + " |")
                    if row_idx == 0:
                        # Add header separator
                        md_lines.append("| " + " | ".join(["---"] * len(row_data)) + " |")
            else:
                # Fallback: just show text content with table indicator
                md_lines.append(f"**[Table {rows}x{cols}]**")
                md_lines.append(text)
            md_lines.append("")  # Empty line after table
        else:
            # Regular paragraph - just add the text
            md_lines.append(text)

    return "\n".join(md_lines)

def docx_find_paragraphs(
    query: str,
    max_results: int = 10,
    return_context: bool = False,
    case_sensitive: bool = False,
    context_span: int = 0
) -> str:
    """
    Find paragraphs containing specific text with result limiting.

    Searches through all paragraphs in the document and returns those containing
    the query text. Limits results to reduce token usage.

    Typical Use Cases:
        - Find placeholders to replace (e.g., "{{NAME}}")
        - Locate paragraphs for modification
        - Search document content programmatically

    Args:        query (str): Text to search for (case-insensitive substring match).
        max_results (int, optional): Maximum results to return. Defaults to 10.
        return_context (bool, optional): Include surrounding context. Defaults to False.

    Returns:
        str: JSON array of objects with paragraph IDs and text:
            [{"id": "para_xxx", "text": "paragraph content"}, ...]

    Raises:
        ValueError: If session_id is invalid or session has expired.

    Examples:
        Find placeholders (limited):
        >>> matches = docx_find_paragraphs("{{NAME}}", max_results=5)

        Find all matches:
        >>> matches = docx_find_paragraphs("important", max_results=999)

    Notes:
        - Search is case-insensitive
        - Results are limited to max_results to save tokens
        - Paragraphs are automatically registered for subsequent operations
        - Empty paragraphs are not searched

    See Also:
        - docx_update_paragraph_text: Modify found paragraphs
        - docx_quick_edit: Find and edit in one step
        - docx_replace_text: Replace text globally
    """
    session, error = get_active_session()
    if error:
        return error

    logger.debug(f"docx_find_paragraphs called: session_id={session.session_id}, query='{query}', max={max_results}")

    matches = []
    paras = list(session.document.paragraphs)
    lowered_query = query if case_sensitive else query.lower()

    for idx, p in enumerate(paras):
        text = p.text
        hay = text if case_sensitive else text.lower()
        if lowered_query in hay:
            p_id = session.register_object(p, "para")
            entry = {"id": p_id, "text": text}

            if return_context and context_span > 0:
                before = []
                after = []
                for i in range(max(0, idx - context_span), idx):
                    before.append(paras[i].text)
                for i in range(idx + 1, min(len(paras), idx + 1 + context_span)):
                    after.append(paras[i].text)
                entry["context_before"] = before
                entry["context_after"] = after

            matches.append(entry)

            if len(matches) >= max_results:
                break

    logger.debug(f"docx_find_paragraphs success: found {len(matches)} matches (limited to {max_results})")

    # Return Markdown format
    if not matches:
        return "No matching paragraphs found."

    md_lines = [f"# Found {len(matches)} matching paragraph(s)\n"]
    for idx, match in enumerate(matches, 1):
        md_lines.append(f"## Match {idx}")
        md_lines.append(f"**ID**: `{match['id']}`")
        md_lines.append(f"**Text**: {match['text']}")

        if return_context and context_span > 0:
            if match.get("context_before"):
                md_lines.append("\n**Context Before**:")
                for ctx in match["context_before"]:
                    md_lines.append(f"> {ctx}")
            if match.get("context_after"):
                md_lines.append("\n**Context After**:")
                for ctx in match["context_after"]:
                    md_lines.append(f"> {ctx}")
        md_lines.append("")

    return "\n".join(md_lines)

def docx_extract_template_structure(
    max_depth: int = None,
    include_content: bool = True,
    max_items_per_type: str = None
) -> str:
    """
    Extract document structure with configurable detail level.

    Analyzes and returns a JSON representation of the document structure,
    including headings, tables, paragraphs, and their properties. Supports
    limiting output to reduce token usage.

    Typical Use Cases:
        - Analyze template structure before filling
        - Understand document layout programmatically
        - Generate documentation from templates
        - Validate template format

    Args:        max_depth (int, optional): Limit nesting depth. None = unlimited.
        include_content (bool, optional): Include text content. Defaults to True.
        max_items_per_type (str, optional): JSON dict limiting items per type.
            Example: '{"headings": 10, "tables": 5, "paragraphs": 0}'

    Returns:
        str: JSON string containing document structure with metadata.

    Raises:
        ValueError: If session_id is invalid or document is empty.

    Examples:
        Full structure:
        >>> structure = docx_extract_template_structure()

        Structure only (no content):
        >>> structure = docx_extract_template_structure(
        ...     session_id, include_content=False
        ... )

        Limited structure:
        >>> structure = docx_extract_template_structure(
        ...     session_id,
        ...     max_items_per_type='{"headings": 10, "tables": 5, "paragraphs": 0}'
        ... )

    Notes:
        - Use max_items_per_type to significantly reduce token usage
        - include_content=False returns only structure metadata
        - Automatically detects table headers (bold or colored background)

    See Also:
        - docx_get_structure_summary: Lightweight alternative
        - docx_read_content: Simple text extraction
    """
    session, error = get_active_session()
    if error:
        return error

    logger.debug(f"docx_extract_template_structure called: session_id={session.session_id}")

    if not session.document:
        logger.error("docx_extract_template_structure failed: Document not found in session")
        raise ValueError("Document not found in session")

    parser = TemplateParser()
    try:
        structure = parser.extract_structure(session.document, session=session)
        doc_structure = structure.get('document_structure', [])

        # Apply filters if specified
        if max_items_per_type or not include_content:
            limits = json.loads(max_items_per_type) if max_items_per_type else {}

            # Count by type
            type_counts = {}
            filtered_structure = []

            for item in doc_structure:
                item_type = item.get('type')
                type_counts[item_type] = type_counts.get(item_type, 0) + 1

                # Check limit (convert 'heading' to 'headings' for lookup)
                limit_key = item_type + 's' if not item_type.endswith('s') else item_type
                limit = limits.get(limit_key, float('inf'))

                if type_counts[item_type] <= limit:
                    # Remove content if requested
                    if not include_content and 'text' in item:
                        item = item.copy()
                        item['text'] = f"[{len(item['text'])} chars]"
                    filtered_structure.append(item)

            structure['document_structure'] = filtered_structure

        element_count = len(structure.get('document_structure', []))
        logger.debug(f"docx_extract_template_structure success: extracted {element_count} elements")

        # Return Markdown format
        md_lines = ["# Document Structure\n"]

        # Summary
        md_lines.append("## Summary")
        md_lines.append(f"- **Total Elements**: {element_count}")
        md_lines.append("")

        # Document structure
        md_lines.append("## Elements\n")
        for idx, item in enumerate(structure.get('document_structure', []), 1):
            item_type = item.get('type', 'unknown')
            md_lines.append(f"### {idx}. {item_type.title()}")

            # Check for both 'id' and 'element_id' for backward compatibility
            element_id = item.get('id') or item.get('element_id')
            if element_id:
                md_lines.append(f"**ID**: `{element_id}`")

            if item_type == 'heading':
                md_lines.append(f"**Level**: {item.get('level', 'N/A')}")
                md_lines.append(f"**Style**: {item.get('style', 'N/A')}")

            if item_type == 'table':
                md_lines.append(f"**Rows**: {item.get('rows', 0)}")
                md_lines.append(f"**Columns**: {item.get('cols', 0)}")
                if item.get('has_header'):
                    md_lines.append("**Has Header**: Yes")
                if 'header_row' in item:
                    md_lines.append(f"**Header Row**: {item['header_row']}")
                if 'headers' in item:
                    md_lines.append(f"**Headers**: {', '.join(item['headers'])}")

            if 'text' in item:
                text = item['text']
                if len(text) > 100:
                    text = text[:100] + "..."
                md_lines.append(f"**Text**: {text}")

            md_lines.append("")

        return "\n".join(md_lines)
    except Exception as e:
        logger.error(f"docx_extract_template_structure failed: {e}")
        raise
