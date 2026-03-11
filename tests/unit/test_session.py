import unittest
import os
import time
from docx import Document
from docx_editor_skill.core.session import SessionManager, Session

class TestSessionManager(unittest.TestCase):
    def setUp(self):
        # Reset the singleton for testing
        SessionManager._instance = None
        self.manager = SessionManager(ttl_seconds=1)

    def test_create_session(self):
        session_id = self.manager.create_session()
        self.assertIsNotNone(session_id)
        self.assertIn(session_id, self.manager.sessions)

        session = self.manager.get_session(session_id)
        self.assertIsInstance(session, Session)
        self.assertIsNotNone(session.document)

    def test_session_expiry(self):
        session_id = self.manager.create_session()
        # Store the session object to check last_accessed later
        session = self.manager.sessions[session_id]
        initial_last_accessed = session.last_accessed

        # Wait for TTL to expire
        time.sleep(1.1)

        # Verify time has passed
        time_passed = time.time() - initial_last_accessed
        self.assertGreater(time_passed, self.manager.ttl_seconds,
                          f"Not enough time passed: {time_passed} <= {self.manager.ttl_seconds}")

        # get_session should return None and remove the expired session
        result = self.manager.get_session(session_id)
        self.assertIsNone(result)
        self.assertNotIn(session_id, self.manager.sessions)

    def test_object_registry(self):
        session_id = self.manager.create_session()
        session = self.manager.get_session(session_id)

        # Test registering a dummy object
        dummy_obj = {"type": "paragraph", "text": "hello"}
        obj_id = session.register_object(dummy_obj, "para")

        self.assertTrue(obj_id.startswith("para_"))
        self.assertEqual(session.get_object(obj_id), dummy_obj)
        self.assertIsNone(session.get_object("non_existent"))

    def test_close_session(self):
        session_id = self.manager.create_session()
        result = self.manager.close_session(session_id)
        self.assertTrue(result)
        self.assertIsNone(self.manager.get_session(session_id))

if __name__ == '__main__':
    unittest.main()
