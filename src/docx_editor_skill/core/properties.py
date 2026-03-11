from typing import Any, Dict, Optional
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL

# Alignment map
ALIGNMENT_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "distribute": WD_ALIGN_PARAGRAPH.DISTRIBUTE,
}

VERTICAL_ALIGNMENT_MAP = {
    "top": WD_ALIGN_VERTICAL.TOP,
    "center": WD_ALIGN_VERTICAL.CENTER,
    "bottom": WD_ALIGN_VERTICAL.BOTTOM,
}

def parse_color(color_str: str) -> Optional[RGBColor]:
    """Parse hex color string to RGBColor."""
    if not color_str:
        return None
    color_str = str(color_str).lstrip("#")
    try:
        if len(color_str) == 6:
            r = int(color_str[0:2], 16)
            g = int(color_str[2:4], 16)
            b = int(color_str[4:6], 16)
            return RGBColor(r, g, b)
    except ValueError:
        pass
    return None

def set_font_properties(font: Any, props: Dict[str, Any]):
    """Helper to set font properties with sugar."""
    for key, value in props.items():
        if key == "color" and isinstance(value, str):
            # Sugar: "color": "FF0000" -> font.color.rgb = ...
            rgb = parse_color(value)
            if rgb and hasattr(font, "color"):
                try:
                    font.color.rgb = rgb
                except Exception:
                    pass
            continue

        if key == "size" and isinstance(value, (int, float)):
            font.size = Pt(value)
            continue

        if key == "name":
            font.name = str(value)
            continue

        # Boolean flags like bold, italic, underline, strike, etc.
        if key in ["bold", "italic", "underline", "strike", "double_strike",
                  "all_caps", "small_caps", "shadow", "outline", "rtl", "imprint", "emboss"]:
            if hasattr(font, key):
                setattr(font, key, bool(value))
            continue

        # Fallback
        if hasattr(font, key):
            try:
                setattr(font, key, value)
            except Exception:
                pass

def set_paragraph_format_properties(fmt: Any, props: Dict[str, Any]):
    """Helper for paragraph format."""
    for key, value in props.items():
        if key == "alignment" and isinstance(value, str):
            if value in ALIGNMENT_MAP:
                fmt.alignment = ALIGNMENT_MAP[value]
            continue

        # Dimension properties - assume Pt if just number
        if key in ["space_before", "space_after", "left_indent", "right_indent", "first_line_indent"]:
            if isinstance(value, (int, float)):
                setattr(fmt, key, Pt(value))
            continue

        if hasattr(fmt, key):
            try:
                setattr(fmt, key, value)
            except Exception:
                pass

def set_properties(obj: Any, properties: Dict[str, Any]):
    """
    Main entry point to set properties on a docx object.
    Supports 'font', 'paragraph_format', and direct attribute setting.
    """

    # 1. Handle "font" block
    if "font" in properties and hasattr(obj, "font"):
        set_font_properties(obj.font, properties["font"])

    # 2. Handle "paragraph_format" block
    if "paragraph_format" in properties and hasattr(obj, "paragraph_format"):
        set_paragraph_format_properties(obj.paragraph_format, properties["paragraph_format"])

    # 3. Handle "table_style" (Table object)
    if "table_style" in properties and hasattr(obj, "style"):
        # obj.style can be set to a style name or style object
        try:
            obj.style = properties["table_style"]
        except Exception:
            pass

    # 4. Handle direct properties
    for key, value in properties.items():
        if key in ["font", "paragraph_format", "table_style"]:
            continue

        # Alignment shortcut for paragraph or objects with alignment
        if key == "alignment" and isinstance(value, str):
            val = ALIGNMENT_MAP.get(value)
            if val is not None:
                if hasattr(obj, "alignment"):
                     obj.alignment = val
                elif hasattr(obj, "paragraph_format"):
                     obj.paragraph_format.alignment = val
            continue

        # Vertical alignment (Cell)
        if key == "vertical_alignment" and isinstance(value, str):
            val = VERTICAL_ALIGNMENT_MAP.get(value)
            if val is not None and hasattr(obj, "vertical_alignment"):
                obj.vertical_alignment = val
            continue

        # Width/Height (generic, assume Pt if number)
        if key in ["width", "height"] and isinstance(value, (int, float)):
             if hasattr(obj, key):
                 setattr(obj, key, Pt(value))
             continue

        # Simple attribute set
        if hasattr(obj, key):
            try:
                setattr(obj, key, value)
            except Exception:
                pass
