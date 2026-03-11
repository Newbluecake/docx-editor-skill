import unittest
import os
import tempfile
from docx_editor_skill.core.session import session_manager
from docx_editor_skill.tools import docx_save

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
from tests.helpers.session_helpers import setup_active_session, teardown_active_session

class TestServerLifecycle(unittest.TestCase):
    def setUp(self):
        # Clear sessions before each test
        session_manager.sessions.clear()

    def test_create_session(self):
        setup_active_session()
        try:
            from docx_editor_skill.core.global_state import global_state
            session_id = global_state.active_session_id
            self.assertIsNotNone(session_id)
            self.assertIn(session_id, session_manager.sessions)
        finally:
            teardown_active_session()

    def test_save_document(self):
        setup_active_session()
        try:
            with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
                tmp_path = tmp.name

            try:
                result = docx_save(tmp_path)
                self.assertTrue(is_success(result))
                self.assertTrue(os.path.exists(tmp_path))
                # Basic check if it's a zip file (docx format)
                with open(tmp_path, 'rb') as f:
                    header = f.read(4)
                    self.assertEqual(header, b'PK\x03\x04')
            finally:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
        finally:
            teardown_active_session()

    def test_save_invalid_session(self):
        # No active session, should fail
        result = docx_save("test.docx")
        self.assertTrue(is_error(result))

if __name__ == '__main__':
    unittest.main()
