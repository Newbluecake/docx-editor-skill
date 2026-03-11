"""Helper functions for extracting data from Markdown responses.

These helpers are used across all test files to extract structured data
from Markdown-formatted tool responses.
"""

import re
import json
import ast
from typing import Optional, Dict, Any


def extract_session_id(response: str) -> Optional[str]:
    """Extract session_id from Markdown response.

    Args:
        response: Markdown response string

    Returns:
        Session ID string or None if not found

    Example:
        >>> response = "**Session ID**: abc123\\n..."
        >>> extract_session_id(response)
        'abc123'
    """
    # Try to extract from Markdown format: **Session ID**: xxx
    match = re.search(r'\*\*Session ID\*\*:\s*(\S+)', response)
    if match:
        return match.group(1)

    # Legacy/Fallback: try **Session Id** (case insensitive)
    match = re.search(r'\*\*Session Id\*\*:\s*(\S+)', response, re.IGNORECASE)
    if match:
        return match.group(1)

    # Fallback: return as-is if it looks like a session ID (short alphanumeric string)
    if isinstance(response, str) and len(response) < 100 and '\n' not in response:
        return response.strip()

    return None


def extract_element_id(response: str) -> Optional[str]:
    """Extract element_id from Markdown response.

    Args:
        response: Markdown response string

    Returns:
        Element ID string or None if not found

    Example:
        >>> response = "**Element ID**: para_abc123\\n..."
        >>> extract_element_id(response)
        'para_abc123'
    """
    # Try to extract from Markdown format: **Element ID**: para_xxx
    match = re.search(r'\*\*Element ID\*\*:\s*(\w+)', response)
    if match:
        return match.group(1)

    # Fallback: try JSON format (legacy compatibility)
    if isinstance(response, str) and response.lstrip().startswith(('{', '[')):
        try:
            data = json.loads(response)
            if isinstance(data, dict) and "data" in data and "element_id" in data["data"]:
                return data["data"]["element_id"]
        except (json.JSONDecodeError, KeyError, TypeError):
            pass

    return response if isinstance(response, str) and response.startswith(('para_', 'run_', 'table_', 'cell_')) else None


def extract_status(response: str) -> str:
    """Extract status from Markdown response.

    Args:
        response: Markdown response string

    Returns:
        "success" or "error"

    Example:
        >>> response = "**Status**: ✅ Success\\n..."
        >>> extract_status(response)
        'success'
    """
    # Try Markdown format
    if '✅ Success' in response or '**Status**: success' in response.lower():
        return 'success'
    if '❌ Error' in response or '**Status**: error' in response.lower():
        return 'error'

    # Fallback: try JSON format only if it looks like JSON
    if isinstance(response, str) and response.lstrip().startswith(('{', '[')):
        try:
            data = json.loads(response)
            if isinstance(data, dict) and "status" in data:
                return data["status"]
        except (json.JSONDecodeError, KeyError, TypeError):
            pass

    return 'unknown'


def _snake_to_title_case(snake_str: str) -> str:
    """Convert snake_case to Title Case, preserving acronyms.

    Args:
        snake_str: Snake case string (e.g., "element_id", "new_row_count")

    Returns:
        Title case string with acronyms preserved (e.g., "Element ID", "New Row Count")
    """
    # Special cases for acronyms that should stay uppercase
    acronyms = {'id': 'ID', 'url': 'URL', 'html': 'HTML', 'xml': 'XML', 'api': 'API'}

    words = snake_str.split('_')
    title_words = []
    for word in words:
        if word.lower() in acronyms:
            title_words.append(acronyms[word.lower()])
        else:
            title_words.append(word.capitalize())

    return ' '.join(title_words)


def extract_metadata_field(response: str, field_name: str) -> Optional[Any]:
    """Extract a specific metadata field from Markdown response.

    Args:
        response: Markdown response string
        field_name: Field name to extract (e.g., "new_row_count", "inserted_at", "element_id")

    Returns:
        Field value or None if not found

    Example:
        >>> response = "**New Row Count**: 4\\n..."
        >>> extract_metadata_field(response, "new_row_count")
        '4'
        >>> response = "**Element ID**: para_123\\n..."
        >>> extract_metadata_field(response, "element_id")
        'para_123'
    """
    # Convert snake_case to Title Case for Markdown format
    display_name = _snake_to_title_case(field_name)

    # Try to extract from Markdown format: **Field Name**: value
    pattern = rf'\*\*{re.escape(display_name)}\*\*:\s*(.+?)(?:\n|$)'
    match = re.search(pattern, response)
    if match:
        value = match.group(1).strip()
        # Try to convert to appropriate type
        if value.lower() in ('true', 'false'):
            return value.lower() == 'true'
        try:
            return int(value)
        except ValueError:
            try:
                return float(value)
            except ValueError:
                return value

    # Fallback: try JSON format
    try:
        data = json.loads(response)
        if isinstance(data, dict) and "data" in data and field_name in data["data"]:
            return data["data"][field_name]
    except (json.JSONDecodeError, KeyError, TypeError):
        pass

    return None


def extract_all_metadata(response: str) -> Dict[str, Any]:
    """Extract all metadata fields from Markdown response.

    Args:
        response: Markdown response string

    Returns:
        Dictionary of all metadata fields

    Example:
        >>> response = "**Element ID**: para_123\\n**New Row Count**: 4\\n..."
        >>> extract_all_metadata(response)
        {'element_id': 'para_123', 'new_row_count': 4}
    """
    metadata = {}

    # Extract common fields
    element_id = extract_element_id(response)
    if element_id:
        metadata['element_id'] = element_id

    session_id = extract_session_id(response)
    if session_id:
        metadata['session_id'] = session_id

    # Extract all **Field**: value patterns
    pattern = r'\*\*([^*]+)\*\*:\s*(.+?)(?:\n|$)'
    for match in re.finditer(pattern, response):
        field_name = match.group(1).strip()
        value = match.group(2).strip()

        # Convert field name to snake_case
        key = field_name.lower().replace(' ', '_')

        # Skip already extracted fields
        if key in ('element_id', 'session_id', 'status', 'operation'):
            continue

        # Try to convert to appropriate type
        if value.lower() in ('true', 'false'):
            metadata[key] = value.lower() == 'true'
        else:
            # Try JSON/literal for complex types
            if (value.startswith('{') and value.endswith('}')) or (value.startswith('[') and value.endswith(']')):
                try:
                    metadata[key] = json.loads(value)
                    continue
                except Exception:
                    try:
                        metadata[key] = ast.literal_eval(value)
                        continue
                    except Exception:
                        pass
            try:
                metadata[key] = int(value)
            except ValueError:
                try:
                    metadata[key] = float(value)
                except ValueError:
                    metadata[key] = value

    return metadata


def is_success(response: str) -> bool:
    """Check if response indicates success.

    Args:
        response: Markdown response string

    Returns:
        True if success, False otherwise
    """
    return extract_status(response) == 'success'


def is_error(response: str) -> bool:
    """Check if response indicates error.

    Args:
        response: Markdown response string

    Returns:
        True if error, False otherwise
    """
    return extract_status(response) == 'error'


def extract_error_message(response: str) -> Optional[str]:
    """Extract error message from Markdown response.

    Args:
        response: Markdown response string

    Returns:
        Error message or None if not found
    """
    # Try Markdown format: **Message**: error text
    match = re.search(r'\*\*Message\*\*:\s*(.+?)(?:\n|$)', response)
    if match:
        return match.group(1).strip()

    # Fallback: try JSON format if it looks like JSON
    if isinstance(response, str) and response.lstrip().startswith(('{', '[')):
        try:
            data = json.loads(response)
            if isinstance(data, dict) and "message" in data:
                return data["message"]
        except (json.JSONDecodeError, KeyError, TypeError):
            pass

    return None


def has_metadata_field(response: str, field_name: str) -> bool:
    """Check if a metadata field exists in Markdown response.

    Args:
        response: Markdown response string
        field_name: Field name to check

    Returns:
        True if field exists, False otherwise
    """
    return extract_metadata_field(response, field_name) is not None


def extract_list_items(response: str, list_marker: str = "-") -> list:
    """Extract list items from Markdown response.

    Args:
        response: Markdown response string
        list_marker: List marker character (default: "-")

    Returns:
        List of extracted items

    Example:
        >>> response = "**Modified Paragraph IDs**:\\n- `para_123`\\n- `para_456`"
        >>> extract_list_items(response)
        ['para_123', 'para_456']
    """
    items = []
    # Match list items with optional code backticks
    pattern = rf'^{re.escape(list_marker)}\s*`?([^`\n]+?)`?\s*$'
    for line in response.split('\n'):
        match = re.match(pattern, line.strip())
        if match:
            items.append(match.group(1).strip())
    return items


def extract_json_from_markdown(response: str) -> Optional[Dict[str, Any]]:
    """Extract JSON data embedded in Markdown response.

    Some composite tools return structured data that needs to be parsed.
    This function extracts key-value pairs from Markdown format.

    Args:
        response: Markdown response string

    Returns:
        Dictionary with extracted data or None

    Example:
        >>> response = "# Quick Edit Result\\n**Modified Count**: 2\\n**Modified Paragraph IDs**:\\n- para_1\\n- para_2"
        >>> extract_json_from_markdown(response)
        {'modified_count': 2, 'paragraph_ids': ['para_1', 'para_2']}
    """
    result = {}

    # Extract numeric fields
    for field in ['modified_count', 'formatted_count', 'rows_filled', 'rows_added']:
        value = extract_metadata_field(response, field)
        if value is not None:
            result[field] = int(value) if isinstance(value, str) and value.isdigit() else value

    # Extract list fields
    if 'Modified Paragraph IDs' in response or 'Paragraph IDs' in response:
        result['paragraph_ids'] = extract_list_items(response)

    # Extract structure summary data (from ## Summary section)
    if '## Summary' in response:
        summary_section = {}
        # Extract Total Headings, Total Tables, Total Paragraphs from list items
        total_headings_match = re.search(r'-\s*\*\*Total Headings\*\*:\s*(\d+)', response)
        total_tables_match = re.search(r'-\s*\*\*Total Tables\*\*:\s*(\d+)', response)
        total_paragraphs_match = re.search(r'-\s*\*\*Total Paragraphs\*\*:\s*(\d+)', response)

        if total_headings_match:
            summary_section['total_headings'] = int(total_headings_match.group(1))
        if total_tables_match:
            summary_section['total_tables'] = int(total_tables_match.group(1))
        if total_paragraphs_match:
            summary_section['total_paragraphs'] = int(total_paragraphs_match.group(1))

        if summary_section:
            result['summary'] = summary_section

    # Extract headings, tables, paragraphs sections
    if '## Headings' in response:
        result['headings'] = _extract_structure_section(response, 'Headings')
    if '## Tables' in response:
        result['tables'] = _extract_structure_section(response, 'Tables')
    if '## Paragraphs' in response:
        result['paragraphs'] = _extract_structure_section(response, 'Paragraphs')

    return result if result else None


def _extract_structure_section(response: str, section_name: str) -> list:
    """Extract a structure section (headings, tables, paragraphs) from Markdown.

    Args:
        response: Markdown response string
        section_name: Section name to extract

    Returns:
        List of items in the section
    """
    items = []
    # Match section, but stop at next major section (Summary, Headings, Tables, Paragraphs) or end
    section_pattern = rf'## {section_name}\s*\n(.*?)(?=\n## (?:Summary|Headings|Tables|Paragraphs)\s*\n|\Z)'
    match = re.search(section_pattern, response, re.DOTALL)

    if match:
        section_content = match.group(1)
        # Extract element IDs from the section - match both formats
        # Format 1: *ID: `xxx`*
        # Format 2: - **ID**: `xxx`
        id_pattern = r'(?:\*ID:\s*`([^`]+)`\*|-\s*\*\*ID\*\*:\s*`([^`]+)`)'
        for id_match in re.finditer(id_pattern, section_content):
            element_id = id_match.group(1) or id_match.group(2)
            if element_id:
                items.append({'element_id': element_id})

    return items


def extract_find_paragraphs_results(response: str) -> list:
    """Extract paragraph search results from docx_find_paragraphs Markdown response.

    Args:
        response: Markdown response from docx_find_paragraphs

    Returns:
        List of dictionaries with 'id', 'text', and optionally 'context_before'/'context_after' keys

    Example:
        >>> response = "# Found 2 matching paragraph(s)\\n\\n## Match 1\\n**ID**: `para_123`\\n**Text**: Hello"
        >>> extract_find_paragraphs_results(response)
        [{'id': 'para_123', 'text': 'Hello'}, ...]
    """
    if "No matching paragraphs found" in response:
        return []

    results = []
    # Split by "## Match N" sections
    match_sections = re.split(r'## Match \d+', response)[1:]  # Skip the header

    for section in match_sections:
        # Extract ID
        id_match = re.search(r'\*\*ID\*\*:\s*`([^`]+)`', section)
        # Extract Text
        text_match = re.search(r'\*\*Text\*\*:\s*(.+?)(?=\n\*\*|\n\n|\Z)', section, re.DOTALL)

        if id_match and text_match:
            result = {
                'id': id_match.group(1).strip(),
                'text': text_match.group(1).strip()
            }

            # Extract context_before if present
            context_before_match = re.search(r'\*\*Context Before\*\*:\s*\n((?:>\s*.+\n?)+)', section)
            if context_before_match:
                context_lines = []
                for line in context_before_match.group(1).split('\n'):
                    line = line.strip()
                    if line.startswith('>'):
                        context_lines.append(line[1:].strip())
                result['context_before'] = context_lines

            # Extract context_after if present
            context_after_match = re.search(r'\*\*Context After\*\*:\s*\n((?:>\s*.+\n?)+)', section)
            if context_after_match:
                context_lines = []
                for line in context_after_match.group(1).split('\n'):
                    line = line.strip()
                    if line.startswith('>'):
                        context_lines.append(line[1:].strip())
                result['context_after'] = context_lines

            results.append(result)

    return results


def extract_template_structure(response: str) -> dict:
    """Extract document structure from docx_extract_template_structure Markdown response.

    Args:
        response: Markdown response from docx_extract_template_structure

    Returns:
        Dictionary with 'document_structure' and 'metadata' keys

    Example:
        >>> response = "# Document Structure\\n\\n## Elements\\n\\n### 1. Heading\\n**ID**: `para_123`"
        >>> extract_template_structure(response)
        {'document_structure': [{'type': 'heading', 'id': 'para_123'}], 'metadata': {...}}
    """
    result = {
        'document_structure': [],
        'metadata': {
            'extracted_at': '',  # Placeholder
            'docx_version': ''   # Placeholder
        }
    }

    # Split by "### N. Type" sections
    element_sections = re.split(r'###\s+\d+\.\s+(\w+)', response)[1:]  # Skip header

    # Process pairs of (type, content)
    for i in range(0, len(element_sections), 2):
        if i + 1 >= len(element_sections):
            break

        element_type = element_sections[i].strip().lower()
        content = element_sections[i + 1]

        element = {'type': element_type}

        # Extract ID
        id_match = re.search(r'\*\*ID\*\*:\s*`([^`]+)`', content)
        if id_match:
            element['id'] = id_match.group(1)

        # Extract Level (for headings)
        level_match = re.search(r'\*\*Level\*\*:\s*(\d+)', content)
        if level_match:
            element['level'] = int(level_match.group(1))

        # Extract Style
        style_match = re.search(r'\*\*Style\*\*:\s*(.+?)(?=\n|$)', content)
        if style_match:
            element['style'] = style_match.group(1).strip()

        # Extract Rows (for tables)
        rows_match = re.search(r'\*\*Rows\*\*:\s*(\d+)', content)
        if rows_match:
            element['rows'] = int(rows_match.group(1))

        # Extract Columns (for tables)
        cols_match = re.search(r'\*\*Columns\*\*:\s*(\d+)', content)
        if cols_match:
            element['cols'] = int(cols_match.group(1))

        # Extract Has Header (for tables)
        if '**Has Header**: Yes' in content:
            element['has_header'] = True

        # Extract Header Row (for tables)
        header_row_match = re.search(r'\*\*Header Row\*\*:\s*(\d+)', content)
        if header_row_match:
            element['header_row'] = int(header_row_match.group(1))

        # Extract Headers (for tables)
        headers_match = re.search(r'\*\*Headers\*\*:\s*(.+?)(?=\n|$)', content)
        if headers_match:
            headers_str = headers_match.group(1).strip()
            element['headers'] = [h.strip() for h in headers_str.split(',')]

        # Extract Text
        text_match = re.search(r'\*\*Text\*\*:\s*(.+?)(?=\n\*\*|\n\n|\Z)', content, re.DOTALL)
        if text_match:
            element['text'] = text_match.group(1).strip()

        result['document_structure'].append(element)

    return result


def extract_element_ids_list(response: str) -> list:
    """Extract list of element IDs from Markdown response.

    Used for tools that return lists of IDs like docx_copy_elements_range.

    Args:
        response: Markdown response containing element IDs

    Returns:
        List of element ID strings

    Example:
        >>> response = "## New Element IDs\\n\\n- **para**: `para_123`\\n- **run**: `run_456`"
        >>> extract_element_ids_list(response)
        ['para_123', 'run_456']
    """
    ids = []
    # Match patterns like: - **type**: `id`
    pattern = r'-\s*\*\*\w+\*\*:\s*`([^`]+)`'
    for match in re.finditer(pattern, response):
        ids.append(match.group(1))
    return ids
