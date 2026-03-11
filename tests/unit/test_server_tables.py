import unittest
import json
from docx_editor_skill.tools import (
    docx_insert_table,
    docx_get_cell,
    docx_insert_paragraph_to_cell,
)
from docx_editor_skill.core.session import session_manager

# Add parent directory to path for helpers import
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from helpers import (
    extract_session_id,
    extract_element_id,
    extract_metadata_field,
    is_success,
    is_error
)

