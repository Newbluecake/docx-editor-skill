from typing import Optional, Any, Iterator, TYPE_CHECKING
from docx.table import Table
from docx.oxml.xmlchemy import BaseOxmlElement

if TYPE_CHECKING:
    from docx.document import Document

class ElementNavigator:
    """
    Helps navigate the underlying XML structure of python-docx objects.
    Provides robust methods to access parents, siblings, and paths.
    """

    @staticmethod
    def get_parent_xml(element_xml: BaseOxmlElement) -> Optional[BaseOxmlElement]:
        """Get the parent XML element."""
        return element_xml.getparent()

    @staticmethod
    def get_docx_parent(element: Any, session_document: Optional['Document'] = None) -> Any:
        """
        Attempt to resolve the high-level python-docx parent object.
        This is tricky because python-docx objects are often transient wrappers around XML.

        Args:
            element: The python-docx object (Paragraph, Table, etc.)
            session_document: Optional Document object to use as root/fallback.

        Returns:
            The parent python-docx object (Document, _Cell, etc.) or None.
        """
        # Try internal attribute first (most reliable for direct children of body)
        if hasattr(element, '_parent') and element._parent:
            return element._parent

        # If we have the XML parent, we might be able to infer context
        if hasattr(element, '_element'):
            parent_xml = element._element.getparent()
            if parent_xml is None:
                return None

            # If parent is body, and we have the document, return document
            if session_document and hasattr(session_document, '_body') and hasattr(session_document._body, '_element'):
                if parent_xml == session_document._body._element:
                    return session_document

            # TODO: Handle Cell parents properly by looking up in session registry?
            # Re-wrapping cells from XML is hard without the Table context.

        return None

    @staticmethod
    def get_next_sibling_xml(element_xml: BaseOxmlElement) -> Optional[BaseOxmlElement]:
        """Get the next sibling XML element."""
        return element_xml.getnext()

    @staticmethod
    def get_prev_sibling_xml(element_xml: BaseOxmlElement) -> Optional[BaseOxmlElement]:
        """Get the previous sibling XML element."""
        return element_xml.getprevious()

    @staticmethod
    def iter_children_xml(parent_xml: BaseOxmlElement) -> Iterator[BaseOxmlElement]:
        """Iterate over child XML elements."""
        return parent_xml.iterchildren()

    @staticmethod
    def get_path(element: Any) -> str:
        """
        Generate a simplified breadcrumb path for the element.
        e.g., "Body/Table[0]/Row[1]/Cell[0]/Paragraph[2]"
        """
        if not hasattr(element, '_element'):
            return "Unknown"

        xml = element._element
        path_parts = []

        current = xml
        while current is not None:
            tag = current.tag.split('}')[-1]  # Remove namespace

            # Calculate index among same-tag siblings
            idx = 0
            prev = current.getprevious()
            while prev is not None:
                if prev.tag == current.tag:
                    idx += 1
                prev = prev.getprevious()

            name = tag
            if tag == 'body':
                name = 'Body'
            elif tag == 'tbl':
                name = 'Table'
            elif tag == 'tr':
                name = 'Row'
            elif tag == 'tc':
                name = 'Cell'
            elif tag == 'p':
                name = 'Para'
            elif tag == 'r':
                name = 'Run'

            path_parts.append(f"{name}[{idx}]")
            current = current.getparent()

            # Stop at document root (usually body's parent is document)
            if current is not None and current.tag.endswith('document'):
                path_parts.append("Document")
                break

        return "/".join(reversed(path_parts))


class ElementManipulator:
    """
    Handles low-level XML insertion and manipulation.
    """

    @staticmethod
    def insert_xml_before(target_xml: BaseOxmlElement, new_xml: BaseOxmlElement):
        """Insert new_xml immediately before target_xml."""
        target_xml.addprevious(new_xml)

    @staticmethod
    def insert_xml_after(target_xml: BaseOxmlElement, new_xml: BaseOxmlElement):
        """Insert new_xml immediately after target_xml."""
        target_xml.addnext(new_xml)

    @staticmethod
    def append_xml_to_parent(parent_xml: BaseOxmlElement, new_xml: BaseOxmlElement):
        """Append new_xml to the end of parent's children."""
        parent_xml.append(new_xml)

    @staticmethod
    def insert_at_index(parent_xml: BaseOxmlElement, new_xml: BaseOxmlElement, index: int):
        """Insert new_xml at specific index in parent."""
        parent_xml.insert(index, new_xml)

    @staticmethod
    def insert_row_at(table: Table, index: int, copy_format_from: Optional[int] = None) -> Any:
        """
        Insert a new row at the specified index in a table.

        Args:
            table: python-docx Table object
            index: Insertion position (0-based), must be in range [0, len(table.rows)]
            copy_format_from: Optional row index to copy formatting from

        Returns:
            The newly inserted _Row object

        Raises:
            IndexError: If index is out of valid range
        """

        num_rows = len(table.rows)
        if index < 0 or index > num_rows:
            raise IndexError(f"Row index {index} out of range (table has {num_rows} rows, valid range: 0-{num_rows})")

        # Create a new row by adding to the end first
        new_row = table.add_row()

        # If we need to insert at a specific position (not at the end)
        if index < num_rows:
            # Remove the row from the end
            table._tbl.remove(new_row._tr)
            # Insert at the desired position
            table._tbl.insert(index, new_row._tr)
            # Re-wrap the row object at the correct position
            new_row = table.rows[index]

        # Copy format if requested
        if copy_format_from is not None:
            if copy_format_from < 0 or copy_format_from >= len(table.rows):
                # If copy_format_from is out of range, skip formatting
                pass
            else:
                from docx_editor_skill.core.format_painter import FormatPainter
                painter = FormatPainter()
                source_row = table.rows[copy_format_from]
                painter.copy_row_format(source_row, new_row)

        return new_row

    @staticmethod
    def insert_col_at(table: Table, index: int, copy_format_from: Optional[int] = None) -> int:
        """
        Insert a new column at the specified index in a table.

        Args:
            table: python-docx Table object
            index: Insertion position (0-based), must be in range [0, len(table.columns)]
            copy_format_from: Optional column index to copy formatting from

        Returns:
            The new column count after insertion

        Raises:
            IndexError: If index is out of valid range
        """
        from docx.oxml.ns import qn
        from docx.shared import Inches
        import copy as copy_module

        # Get current column count from the first row (more reliable than table.columns)
        num_cols = len(table.rows[0].cells) if len(table.rows) > 0 else 0
        if index < 0 or index > num_cols:
            raise IndexError(f"Column index {index} out of range (table has {num_cols} columns, valid range: 0-{num_cols})")

        # Insert a new cell in each row at the specified index
        for row in table.rows:
            # Create a new cell by deep copying an existing cell's XML structure
            # This ensures all namespaces and required elements are present
            if len(row.cells) > 0:
                # Copy the first cell as a template
                template_tc = row.cells[0]._tc
                new_tc = copy_module.deepcopy(template_tc)

                # Clear the content of the new cell (remove all paragraphs except one empty one)
                for p in list(new_tc.findall(qn('w:p'))):
                    new_tc.remove(p)
                # Add one empty paragraph
                new_tc._new_p()
            else:
                # Fallback: create a minimal cell (shouldn't happen in practice)
                from docx.oxml.table import CT_Tc
                new_tc = CT_Tc()
                new_tc._new_p()

            # Insert the cell at the specified index
            row._tr.insert(index, new_tc)

        # Update table grid (column width definitions)
        tbl_grid = table._tbl.tblGrid
        if tbl_grid is not None:
            from docx.oxml.table import CT_TblGridCol
            new_grid_col = CT_TblGridCol()
            # Set default width (1 inch)
            new_grid_col.w = Inches(1.0)
            tbl_grid.insert(index, new_grid_col)

        # Copy format if requested
        if copy_format_from is not None:
            new_col_count = len(table.rows[0].cells) if len(table.rows) > 0 else 0
            if copy_format_from < 0 or copy_format_from >= new_col_count:
                # If copy_format_from is out of range, skip formatting
                pass
            else:
                from docx_editor_skill.core.format_painter import FormatPainter
                painter = FormatPainter()
                painter.copy_col_format(table, copy_format_from, index)

        # Return the new column count
        return len(table.rows[0].cells) if len(table.rows) > 0 else 0

    @staticmethod
    def delete_row(table: Table, index: int) -> None:
        """
        Delete a row from the table at the specified index.

        Args:
            table: python-docx Table object
            index: Row index to delete (0-based)

        Raises:
            IndexError: If index is out of valid range
            ValueError: If attempting to delete the last row
        """
        num_rows = len(table.rows)

        if num_rows == 1:
            raise ValueError("Cannot delete the last row")

        if index < 0 or index >= num_rows:
            raise IndexError(f"Row index {index} out of range (table has {num_rows} rows)")

        # Get the row and remove it from the table
        row = table.rows[index]
        table._tbl.remove(row._tr)

    @staticmethod
    def delete_col(table: Table, index: int) -> int:
        """
        Delete a column from the table at the specified index.

        Args:
            table: python-docx Table object
            index: Column index to delete (0-based)

        Returns:
            The new column count after deletion

        Raises:
            IndexError: If index is out of valid range
            ValueError: If attempting to delete the last column
        """
        # Get current column count from the first row
        num_cols = len(table.rows[0].cells) if len(table.rows) > 0 else 0

        if num_cols == 1:
            raise ValueError("Cannot delete the last column")

        if index < 0 or index >= num_cols:
            raise IndexError(f"Column index {index} out of range (table has {num_cols} columns)")

        # Remove the cell at the specified index from each row
        for row in table.rows:
            if index < len(row.cells):
                cell = row.cells[index]
                row._tr.remove(cell._tc)

        # Update table grid (column width definitions)
        tbl_grid = table._tbl.tblGrid
        if tbl_grid is not None and index < len(tbl_grid.gridCol_lst):
            grid_col = tbl_grid.gridCol_lst[index]
            tbl_grid.remove(grid_col)

        # Return the new column count
        return len(table.rows[0].cells) if len(table.rows) > 0 else 0
