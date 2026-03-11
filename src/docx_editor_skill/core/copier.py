import copy
from docx.table import Table

def clone_table(table: Table) -> Table:
    """
    Deep copy a table using XML manipulation.
    Returns a new Table object attached to the same parent but initially not inserted in the document flow?
    Actually, python-docx doesn't easily let us create a "floating" table wrapper.
    We usually need to append the element to a parent.

    Strategy:
    1. Deep copy the underlying XML element (tbl).
    2. Append it to the parent (document body or cell).
    3. Return the new Table wrapper.
    """
    tbl_element = copy.deepcopy(table._tbl)

    # We assume the table is in the document body for now.
    # If it was nested, table._parent would be the cell/doc.
    parent = table._parent

    # Append to the parent's element
    parent._element.append(tbl_element)

    # Create wrapper
    new_table = Table(tbl_element, parent)

    return new_table
