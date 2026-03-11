from typing import Any, Dict, Optional, Union, Literal
from dataclasses import dataclass, asdict
import json
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docx.table import Table
from docx.shared import RGBColor

@dataclass
class FontProperties:
    name: Optional[str] = None
    size: Optional[int] = None  # Stored as raw value (EMU usually if coming from Length)
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    strike: Optional[bool] = None
    color_rgb: Optional[str] = None  # Hex string like "FF0000"
    highlight_color: Optional[int] = None # Enum value

@dataclass
class ParagraphProperties:
    alignment: Optional[int] = None  # WD_ALIGN_PARAGRAPH enum value
    left_indent: Optional[int] = None
    right_indent: Optional[int] = None
    first_line_indent: Optional[int] = None
    space_before: Optional[int] = None
    space_after: Optional[int] = None
    line_spacing: Optional[float] = None # Can be float or length
    line_spacing_rule: Optional[int] = None
    keep_together: Optional[bool] = None
    keep_with_next: Optional[bool] = None
    page_break_before: Optional[bool] = None
    style_name: Optional[str] = None

@dataclass
class TableProperties:
    style_name: Optional[str] = None
    # Deep table properties (borders, shading) are complex to serialize fully portably
    # without raw XML. For MVP, we'll store basic style.
    # Future: store XML for borders/shading if needed, or breakdown properties.

@dataclass
class FormatTemplate:
    element_type: Literal["paragraph", "run", "table"]
    properties: Dict[str, Any]
    font_properties: Optional[Dict[str, Any]] = None  # Paragraphs also have font props (run style)

class TemplateManager:
    """
    Manages extraction and application of formatting templates.
    """

    def extract_template(self, element: Union[Paragraph, Run, Table]) -> FormatTemplate:
        if isinstance(element, Run):
            return self._extract_run_template(element)
        elif isinstance(element, Paragraph):
            return self._extract_paragraph_template(element)
        elif isinstance(element, Table):
            return self._extract_table_template(element)
        else:
            raise ValueError(f"Unsupported element type: {type(element)}")

    def apply_template(self, element: Union[Paragraph, Run, Table], template: FormatTemplate) -> None:
        if template.element_type == "run" and isinstance(element, Run):
            self._apply_run_template(element, template)
        elif template.element_type == "paragraph" and isinstance(element, Paragraph):
            self._apply_paragraph_template(element, template)
        elif template.element_type == "table" and isinstance(element, Table):
            self._apply_table_template(element, template)
        else:
            # Cross-application (e.g. Paragraph template to Run?) - Not supported in MVP T-002
            # But could be handled like FormatPainter
            raise ValueError(f"Mismatch or unsupported: Apply {template.element_type} template to {type(element)}")

    def to_json(self, template: FormatTemplate) -> str:
        return json.dumps(asdict(template))

    def from_json(self, json_str: str) -> FormatTemplate:
        data = json.loads(json_str)
        return FormatTemplate(**data)

    # --- Internal Extraction Methods ---

    def _extract_font_props(self, font_obj: Any) -> Dict[str, Any]:
        props = {}
        if font_obj.name:
            props['name'] = font_obj.name
        if font_obj.size:
            props['size'] = font_obj.size
        if font_obj.bold is not None:
            props['bold'] = font_obj.bold
        if font_obj.italic is not None:
            props['italic'] = font_obj.italic
        if font_obj.underline is not None:
            props['underline'] = font_obj.underline
        if font_obj.strike is not None:
            props['strike'] = font_obj.strike

        if font_obj.color and font_obj.color.rgb:
            props['color_rgb'] = str(font_obj.color.rgb)

        # Highlight is often on font object in python-docx
        if hasattr(font_obj, 'highlight_color') and font_obj.highlight_color is not None:
             props['highlight_color'] = font_obj.highlight_color

        return props

    def _extract_run_template(self, run: Run) -> FormatTemplate:
        font_props = self._extract_font_props(run.font)
        return FormatTemplate(element_type="run", properties={}, font_properties=font_props)

    def _extract_paragraph_template(self, para: Paragraph) -> FormatTemplate:
        fmt = para.paragraph_format
        props = {}

        if fmt.alignment is not None:
            props['alignment'] = fmt.alignment
        if fmt.left_indent is not None:
            props['left_indent'] = fmt.left_indent
        if fmt.right_indent is not None:
            props['right_indent'] = fmt.right_indent
        if fmt.first_line_indent is not None:
            props['first_line_indent'] = fmt.first_line_indent
        if fmt.space_before is not None:
            props['space_before'] = fmt.space_before
        if fmt.space_after is not None:
            props['space_after'] = fmt.space_after
        if fmt.line_spacing is not None:
            props['line_spacing'] = fmt.line_spacing
        if fmt.line_spacing_rule is not None:
            props['line_spacing_rule'] = fmt.line_spacing_rule

        if fmt.keep_together is not None:
            props['keep_together'] = fmt.keep_together
        if fmt.keep_with_next is not None:
            props['keep_with_next'] = fmt.keep_with_next
        if fmt.page_break_before is not None:
            props['page_break_before'] = fmt.page_break_before

        if para.style:
            props['style_name'] = para.style.name

        return FormatTemplate(element_type="paragraph", properties=props)

    def _extract_table_template(self, table: Table) -> FormatTemplate:
        props = {}
        if table.style:
            props['style_name'] = table.style.name
        return FormatTemplate(element_type="table", properties=props)

    # --- Internal Application Methods ---

    def _apply_font_props(self, font_obj: Any, props: Dict[str, Any]) -> None:
        if 'name' in props:
            font_obj.name = props['name']
        if 'size' in props:
            font_obj.size = props['size']
        if 'bold' in props:
            font_obj.bold = props['bold']
        if 'italic' in props:
            font_obj.italic = props['italic']
        if 'underline' in props:
            font_obj.underline = props['underline']
        if 'strike' in props:
            font_obj.strike = props['strike']

        if 'color_rgb' in props:
            font_obj.color.rgb = RGBColor.from_string(props['color_rgb'])

        if 'highlight_color' in props and hasattr(font_obj, 'highlight_color'):
            font_obj.highlight_color = props['highlight_color']

    def _apply_run_template(self, run: Run, template: FormatTemplate) -> None:
        if template.font_properties:
            self._apply_font_props(run.font, template.font_properties)

    def _apply_paragraph_template(self, para: Paragraph, template: FormatTemplate) -> None:
        fmt = para.paragraph_format
        props = template.properties

        if 'alignment' in props:
            fmt.alignment = props['alignment']
        if 'left_indent' in props:
            fmt.left_indent = props['left_indent']
        if 'right_indent' in props:
            fmt.right_indent = props['right_indent']
        if 'first_line_indent' in props:
            fmt.first_line_indent = props['first_line_indent']
        if 'space_before' in props:
            fmt.space_before = props['space_before']
        if 'space_after' in props:
            fmt.space_after = props['space_after']
        if 'line_spacing' in props:
            fmt.line_spacing = props['line_spacing']
        if 'line_spacing_rule' in props:
            fmt.line_spacing_rule = props['line_spacing_rule']

        if 'keep_together' in props:
            fmt.keep_together = props['keep_together']
        if 'keep_with_next' in props:
            fmt.keep_with_next = props['keep_with_next']
        if 'page_break_before' in props:
            fmt.page_break_before = props['page_break_before']

        if 'style_name' in props:
            try:
                para.style = props['style_name']
            except (ValueError, KeyError):
                # Style might not exist in document, ignore or log warning
                pass

    def _apply_table_template(self, table: Table, template: FormatTemplate) -> None:
        props = template.properties
        if 'style_name' in props:
             try:
                table.style = props['style_name']
             except (ValueError, KeyError):
                pass

        # NOTE: Deep XML structure application (borders/shading) is not implemented in MVP
        # as it requires deepcopying OXML elements which are not easily serializable to JSON
        # without storing raw XML strings.
