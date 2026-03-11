"""ASCII visualization and diff rendering for document elements.

This module provides ASCII-based visualization of document structure and
Git-style diff rendering for content changes.
"""

import logging
from typing import List, Tuple, Optional
from difflib import SequenceMatcher
from docx.text.paragraph import Paragraph
from docx.table import Table

logger = logging.getLogger(__name__)


class DocumentVisualizer:
    """ASCII visualization renderer for document elements."""

    def __init__(self, session):
        """Initialize visualizer with session context.

        Args:
            session: Session object containing document and registry
        """
        self.session = session
        self.max_width = 80  # Maximum line width for content
        self.context_range = 7  # Elements before/after current element

    def render_paragraph(self, paragraph: Paragraph, element_id: str,
                        highlight: bool = False) -> str:
        """Render a paragraph as ASCII box.

        Args:
            paragraph: Paragraph object
            element_id: Element ID
            highlight: Whether to mark as current/new

        Returns:
            ASCII box representation
        """
        # Extract text with formatting
        text = self._extract_text_with_format(paragraph)

        # Truncate if too long
        if len(text) > self.max_width:
            text = self._truncate_text(text, self.max_width)

        # Build title
        title = f"Paragraph ({element_id})"
        if highlight:
            title += " ⭐ CURRENT"

        # Draw box
        return self._draw_box(text, title)

    def render_table(self, table: Table, element_id: str,
                    highlight: bool = False) -> str:
        """Render a table as ASCII grid.

        Args:
            table: Table object
            element_id: Element ID
            highlight: Whether to mark as current/new

        Returns:
            ASCII table representation
        """
        rows = len(table.rows)
        cols = len(table.columns) if table.rows else 0

        # Build title
        title = f"Table ({element_id})"
        if highlight:
            title += " ⭐ CURRENT"

        # Limit table size for display
        max_display_rows = min(rows, 20)
        max_display_cols = min(cols, 10)

        # Extract cell contents
        cell_data = []
        for i in range(max_display_rows):
            row_data = []
            for j in range(max_display_cols):
                try:
                    cell = table.rows[i].cells[j]
                    # Get first paragraph text
                    text = cell.text.strip() if cell.text else "(empty)"
                    # Truncate cell content
                    text = self._truncate_text(text, 20)
                    row_data.append(text)
                except (IndexError, AttributeError):
                    row_data.append("(empty)")
            cell_data.append(row_data)

        # Calculate column widths
        col_widths = []
        for j in range(max_display_cols):
            max_width = max(len(row[j]) for row in cell_data) if cell_data else 10
            col_widths.append(max(max_width, 10))

        # Build table content
        lines = []

        # Title line
        total_width = sum(col_widths) + (max_display_cols - 1) * 3 + 2
        lines.append(f"│ {title.ljust(total_width - 2)} │")

        # Draw table grid
        for i, row in enumerate(cell_data):
            # Draw separator before first row
            if i == 0:
                sep = "├" + "┬".join("─" * (w + 2) for w in col_widths) + "┤"
                lines.append(sep)

            # Draw row content
            row_str = "│"
            for j, cell_text in enumerate(row):
                row_str += f" {cell_text.ljust(col_widths[j])} │"
            lines.append(row_str)

            # Draw separator between rows (except last)
            if i < len(cell_data) - 1:
                sep = "├" + "┼".join("─" * (w + 2) for w in col_widths) + "┤"
                lines.append(sep)

        # Add truncation notice if needed
        if rows > max_display_rows or cols > max_display_cols:
            truncation_msg = f"(showing {max_display_rows}/{rows} rows, {max_display_cols}/{cols} cols)"
            lines.append(f"│ {truncation_msg.ljust(total_width - 2)} │")

        # Wrap in box
        box_width = total_width
        result = ["┌" + "─" * box_width + "┐"]
        result.extend(lines)
        result.append("└" + "─" * box_width + "┘")

        return "\n".join(result)

    def render_context(self, element_id: str,
                      context_range: Optional[int] = None) -> str:
        """Render document context around an element.

        Args:
            element_id: Current element ID
            context_range: Number of elements before/after (default: 7)

        Returns:
            ASCII visualization of context
        """
        if context_range is None:
            context_range = self.context_range

        # Get all top-level elements (paragraphs and tables)
        elements = []
        for element in self.session.document.element.body:
            # Check if it's a paragraph or table
            if element.tag.endswith('p'):
                # Find corresponding Paragraph object
                for para in self.session.document.paragraphs:
                    if para._element == element:
                        para_id = self.session._get_element_id(para, auto_register=True)
                        elements.append(('paragraph', para, para_id))
                        break
            elif element.tag.endswith('tbl'):
                # Find corresponding Table object
                for table in self.session.document.tables:
                    if table._element == element:
                        table_id = self.session._get_element_id(table, auto_register=True)
                        elements.append(('table', table, table_id))
                        break

        # Find current element index
        current_index = -1
        for i, (elem_type, elem_obj, elem_id) in enumerate(elements):
            if elem_id == element_id:
                current_index = i
                break

        if current_index == -1:
            return f"Element {element_id} not found in document"

        # Calculate context range
        start_index = max(0, current_index - context_range)
        end_index = min(len(elements), current_index + context_range + 1)

        # Build context visualization
        lines = []
        lines.append(f"📄 Document Context (showing {end_index - start_index} elements around {element_id})")
        lines.append("")

        # Add ellipsis if not at start
        if start_index > 0:
            lines.append(f"  ... ({start_index} more elements above) ...")
            lines.append("")

        # Render elements in range
        for i in range(start_index, end_index):
            elem_type, elem_obj, elem_id = elements[i]

            # Add cursor marker before current element
            if i == current_index:
                lines.append(">>> [CURSOR] <<<")
                lines.append("")

            # Render element
            highlight = (i == current_index)
            if elem_type == 'paragraph':
                rendered = self.render_paragraph(elem_obj, elem_id, highlight)
            else:  # table
                rendered = self.render_table(elem_obj, elem_id, highlight)

            # Indent each line
            for line in rendered.split('\n'):
                lines.append(f"  {line}")
            lines.append("")

        # Add ellipsis if not at end
        if end_index < len(elements):
            lines.append(f"  ... ({len(elements) - end_index} more elements below) ...")

        return "\n".join(lines)

    def render_image(self, image_path: str, element_id: str) -> str:
        """Render image placeholder.

        Args:
            image_path: Path to image file
            element_id: Element ID

        Returns:
            Image placeholder string
        """
        import os
        filename = os.path.basename(image_path)
        return f"[IMG: {filename}]"

    def render_cursor(self) -> str:
        """Render cursor marker.

        Returns:
            Cursor marker string
        """
        return ">>> [CURSOR] <<<"

    def _extract_text_with_format(self, paragraph: Paragraph) -> str:
        """Extract paragraph text with Markdown formatting.

        Converts:
        - Bold runs → **text**
        - Italic runs → *text*

        Args:
            paragraph: Paragraph object

        Returns:
            Formatted text string
        """
        result = []
        for run in paragraph.runs:
            text = run.text
            if not text:
                continue

            # Apply formatting
            if run.bold and run.italic:
                text = f"***{text}***"
            elif run.bold:
                text = f"**{text}**"
            elif run.italic:
                text = f"*{text}*"

            result.append(text)

        return "".join(result) if result else paragraph.text

    def _truncate_text(self, text: str, max_length: int = 80) -> str:
        """Truncate text and add ellipsis if needed.

        Args:
            text: Text to truncate
            max_length: Maximum length

        Returns:
            Truncated text
        """
        if len(text) <= max_length:
            return text
        return text[:max_length - 3] + "..."

    def _draw_box(self, content: str, title: str) -> str:
        """Draw ASCII box around content.

        Args:
            content: Content to wrap
            title: Box title

        Returns:
            ASCII box string
        """
        lines = content.split('\n') if content else ["(empty)"]

        # Calculate max width
        max_width = max(len(title), max(len(line) for line in lines))
        max_width = min(max_width, self.max_width)

        # Build box
        result = []
        result.append('┌' + '─' * (max_width + 2) + '┐')
        result.append('│ ' + title.ljust(max_width) + ' │')
        result.append('├' + '─' * (max_width + 2) + '┤')

        for line in lines:
            # Truncate line if needed
            if len(line) > max_width:
                line = line[:max_width - 3] + "..."
            result.append('│ ' + line.ljust(max_width) + ' │')

        result.append('└' + '─' * (max_width + 2) + '┘')

        return '\n'.join(result)


class DiffRenderer:
    """Git-style diff renderer for content changes."""

    def render_diff(self, old_content: str, new_content: str,
                   element_id: str, element_type: str = "Paragraph") -> str:
        """Render Git-style diff.

        Args:
            old_content: Original content
            new_content: Modified content
            element_id: Element ID
            element_type: Type of element (Paragraph, Table, etc.)

        Returns:
            Markdown-formatted diff
        """
        # Split into lines
        old_lines = old_content.split('\n') if old_content else []
        new_lines = new_content.split('\n') if new_content else []

        # Compute diff
        diff_lines = self._compute_line_diff(old_lines, new_lines)

        # Build diff output
        lines = []
        lines.append("🔄 Changes")
        lines.append("")

        # Draw box with diff
        title = f"{element_type} ({element_id})"
        max_width = 80

        lines.append('┌' + '─' * (max_width + 2) + '┐')
        lines.append('│ ' + title.ljust(max_width) + ' │')
        lines.append('├' + '─' * (max_width + 2) + '┤')

        for prefix, line in diff_lines:
            # Truncate line if needed
            display_line = line[:max_width - 2] if len(line) > max_width - 2 else line
            lines.append(f'{prefix} │ {display_line.ljust(max_width - 2)} │')

        lines.append('└' + '─' * (max_width + 2) + '┘')

        return '\n'.join(lines)

    def _compute_line_diff(self, old_lines: List[str],
                          new_lines: List[str]) -> List[Tuple[str, str]]:
        """Compute line-by-line diff.

        Args:
            old_lines: Original lines
            new_lines: Modified lines

        Returns:
            List of (prefix, line) tuples where prefix is ' ', '-', or '+'
        """
        matcher = SequenceMatcher(None, old_lines, new_lines)
        result = []

        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                # Unchanged lines
                for line in old_lines[i1:i2]:
                    result.append((' ', line))
            elif tag == 'delete':
                # Deleted lines
                for line in old_lines[i1:i2]:
                    result.append(('-', line))
            elif tag == 'insert':
                # Inserted lines
                for line in new_lines[j1:j2]:
                    result.append(('+', line))
            elif tag == 'replace':
                # Replaced lines (show as delete + insert)
                for line in old_lines[i1:i2]:
                    result.append(('-', line))
                for line in new_lines[j1:j2]:
                    result.append(('+', line))

        return result
