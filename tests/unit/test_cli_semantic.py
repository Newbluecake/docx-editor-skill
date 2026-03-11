"""Unit tests for semantic CLI commands (edit, insert, format, copy)."""

import json
import pytest
from docx import Document

from docx_editor_skill.cli import (
    _is_element_id,
    _is_image_path,
    _parse_heading,
    _is_json_dict,
    _resolve_position,
    _has_format_args,
)


# ---------------------------------------------------------------------------
# Helper function tests
# ---------------------------------------------------------------------------

def test_is_element_id():
    assert _is_element_id("para_001")
    assert _is_element_id("run_123")
    assert _is_element_id("table_999")
    assert _is_element_id("cell_042")
    assert not _is_element_id("para_1")  # Wrong format
    assert not _is_element_id("para_0001")  # Too many digits
    assert not _is_element_id("foo_001")  # Wrong prefix
    assert not _is_element_id("para001")  # Missing underscore


def test_is_image_path():
    assert _is_image_path("image.png")
    assert _is_image_path("photo.jpg")
    assert _is_image_path("chart.JPEG")
    assert _is_image_path("/path/to/file.gif")
    assert _is_image_path("./relative.bmp")
    assert _is_image_path("file.tiff")
    assert not _is_image_path("document.docx")
    assert not _is_image_path("text.txt")
    assert not _is_image_path("no_extension")


def test_parse_heading():
    assert _parse_heading("# Title") == (1, "Title")
    assert _parse_heading("## Subtitle") == (2, "Subtitle")
    assert _parse_heading("### Section") == (3, "Section")
    assert _parse_heading("###### Level 6") == (6, "Level 6")
    assert _parse_heading("# Title with spaces  ") == (1, "Title with spaces  ")
    assert _parse_heading("Not a heading") is None
    assert _parse_heading("##No space") is None
    assert _parse_heading("####### Too many") is None


def test_is_json_dict():
    assert _is_json_dict('{"key": "value"}')
    assert _is_json_dict('{"a": 1, "b": 2}')
    assert not _is_json_dict('["array"]')
    assert not _is_json_dict('"string"')
    assert not _is_json_dict('123')
    assert not _is_json_dict('not json')
    assert not _is_json_dict('')


def test_resolve_position():
    class Args:
        pass

    args = Args()
    args.after = "para_001"
    assert _resolve_position(args) == "after:para_001"

    args = Args()
    args.before = "table_002"
    assert _resolve_position(args) == "before:table_002"

    args = Args()
    args.inside = "cell_003"
    assert _resolve_position(args) == "inside:cell_003"

    args = Args()
    # No position flags → default
    assert _resolve_position(args) == "end:document_body"


def test_has_format_args():
    class Args:
        pass

    args = Args()
    args.bold = True
    args.italic = None
    args.size = None
    args.color = None
    assert _has_format_args(args)

    args = Args()
    args.bold = None
    args.italic = None
    args.size = 14.0
    args.color = None
    assert _has_format_args(args)

    args = Args()
    args.bold = None
    args.italic = None
    args.size = None
    args.color = None
    assert not _has_format_args(args)


# ---------------------------------------------------------------------------
# Integration tests (require CLI session setup)
# ---------------------------------------------------------------------------
# These would test the actual command handlers with a real document.
# For now, we've validated the routing logic with unit tests above.
# Full integration tests would be added in tests/e2e/.
