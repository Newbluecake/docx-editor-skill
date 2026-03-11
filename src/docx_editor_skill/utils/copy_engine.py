from copy import deepcopy
from typing import Optional, Any, Union, List
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.xmlchemy import BaseOxmlElement
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl


class CopyEngine:
    """
    Handles deep copying of docx elements (Paragraphs, Tables) preserving full fidelity
    via XML manipulation.
    """

    def copy_element(self, element: Union[Paragraph, Table]) -> BaseOxmlElement:
        """
        Creates a deep copy of the element's underlying XML.
        Returns the orphaned XML element.
        """
        if not hasattr(element, '_element'):
            raise ValueError(f"Cannot copy object of type {type(element)}: missing _element")

        return deepcopy(element._element)

    def get_elements_between(self, start_el: Union[Paragraph, Table], end_el: Union[Paragraph, Table]) -> List[Union[Paragraph, Table]]:
        """
        Retrieves all elements between start_el and end_el (inclusive).
        Elements must share the same parent.
        """
        if not hasattr(start_el, '_element') or not hasattr(end_el, '_element'):
             raise ValueError("Elements must have underlying _element")

        start_xml = start_el._element
        end_xml = end_el._element
        parent = start_xml.getparent()

        if end_xml.getparent() != parent:
            raise ValueError("Start and end elements must share the same parent")

        # Get indices
        try:
            start_idx = parent.index(start_xml)
            end_idx = parent.index(end_xml)
        except ValueError:
             raise ValueError("Elements not found in parent")

        if start_idx > end_idx:
            raise ValueError("Start element must appear before end element")

        # Collect elements in range
        elements = []
        # Slice includes end_idx so use end_idx + 1
        for i in range(start_idx, end_idx + 1):
            child = parent[i]
            # Wrap child back to docx object
            # We need to pass the docx Parent object (e.g. Document) to the constructor
            # But the 'parent' here is an lxml element.
            # We can use start_el._parent (python-docx attribute) as the parent for the new wrappers
            docx_parent = getattr(start_el, '_parent', None)

            # Filter for supported types (p and tbl)
            if isinstance(child, (CT_P, CT_Tbl)):
                 elements.append(self._wrap_element(child, docx_parent))
            # Ignore other elements (e.g. sectPr, comments, etc.) for now as they aren't registered objects

        return elements

    def copy_range(self, start_el: Union[Paragraph, Table], end_el: Union[Paragraph, Table],
                   target_parent: Any, ref_element_xml: Optional[BaseOxmlElement] = None) -> List[Any]:
        """
        Copies a range of elements and inserts them into target_parent.

        Args:
            start_el: Start of range (inclusive)
            end_el: End of range (inclusive)
            target_parent: The docx object to insert into (Document, Cell, etc.)
            ref_element_xml: Optional insertion point (insert after this)

        Returns:
            List of new python-docx objects
        """
        source_elements = self.get_elements_between(start_el, end_el)
        new_objects = []

        # Maintain insertion order.
        # If ref_element_xml is provided, we insert sequentially after it.
        # For the first element, insert after ref_element_xml.
        # For subsequent elements, insert after the PREVIOUSLY inserted element.

        current_ref = ref_element_xml

        for el in source_elements:
            # 1. Copy XML
            new_xml = self.copy_element(el)

            # 2. Insert
            # If current_ref is None, it appends to the end of container.
            # However, if we are appending a sequence [A, B, C] and current_ref is None:
            # Append A -> [..., A]
            # Append B -> [..., A, B]
            # So simple append works if ref is None.

            # If current_ref is NOT None (inserting in middle):
            # Insert A after Ref -> [..., Ref, A, ...] -> current_ref becomes A
            # Insert B after A   -> [..., Ref, A, B, ...]

            new_obj = self.insert_element_after(target_parent, new_xml, current_ref)
            new_objects.append(new_obj)

            # Update reference for next iteration so we chain them
            if current_ref is not None:
                current_ref = new_xml

        return new_objects

    def insert_element_after(self, parent: Any, new_element_xml: BaseOxmlElement, ref_element_xml: Optional[BaseOxmlElement] = None) -> Any:
        """
        Inserts the new XML element into the parent's XML container.

        Args:
            parent: The python-docx object serving as parent (Document, Cell, etc.)
            new_element_xml: The copied XML element
            ref_element_xml: The element to insert after. If None, appends to end.

        Returns:
            The wrapped python-docx object (Paragraph or Table)
        """
        # Determine container element (usually body or cell element)
        # parent usually has ._body (if Document) or is a Cell
        if hasattr(parent, '_body') and hasattr(parent._body, '_body'):
            # This handles docx.Document object which wraps a Body object
            container = parent._body._body
        elif hasattr(parent, '_element'):
            # Cell or other container wrappers
            container = parent._element
        else:
            # Fallback: maybe passed the body element directly?
            raise ValueError(f"Unsupported parent type: {type(parent)}")

        # Insert logic
        if ref_element_xml is not None:
            ref_element_xml.addnext(new_element_xml)
        else:
            container.append(new_element_xml)

        # Wrap back to python-docx object
        return self._wrap_element(new_element_xml, parent)

    def _wrap_element(self, element_xml: BaseOxmlElement, parent: Any) -> Any:
        """
        Wraps an XML element into its corresponding python-docx object.
        """
        if isinstance(element_xml, CT_P):
            return Paragraph(element_xml, parent)
        elif isinstance(element_xml, CT_Tbl):
            return Table(element_xml, parent)
        else:
            # Try to infer from tag if instance check fails (unlikely with python-docx classes but safer)
            tag = element_xml.tag
            if tag.endswith('p'):
                return Paragraph(element_xml, parent)
            elif tag.endswith('tbl'):
                return Table(element_xml, parent)

            raise ValueError(f"Unsupported XML element type for wrapping: {type(element_xml)} / {tag}")
