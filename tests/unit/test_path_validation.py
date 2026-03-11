import os
import pytest
import tempfile
from docx_editor_skill.core.session import SessionManager

def test_create_session_valid_new_file():
    """Test creating a session for a new file in an existing directory."""
    manager = SessionManager()
    with tempfile.TemporaryDirectory() as tmpdir:
        new_file_path = os.path.join(tmpdir, "new_doc.docx")

        # Should succeed
        session_id = manager.create_session(file_path=new_file_path)
        assert session_id is not None

        # Cleanup
        manager.close_session(session_id)

def test_create_session_invalid_parent_dir():
    """Test creating a session for a file where the parent directory doesn't exist."""
    manager = SessionManager()

    # Construct a path that definitely shouldn't exist
    # Use a GUID to avoid collision with any weird real path
    bad_path = os.path.join(tempfile.gettempdir(), "nonexistent_dir_xyz123", "doc.docx")

    # Should raise ValueError
    with pytest.raises(ValueError) as excinfo:
        manager.create_session(file_path=bad_path)

    assert "Parent directory does not exist" in str(excinfo.value)

def test_create_session_existing_file():
    """Test creating a session for a file that already exists."""
    manager = SessionManager()
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp_path = tmp.name

    try:
        # Should succeed (loading existing file)
        # Note: python-docx might fail if the temp file is empty/invalid docx,
        # but here we are mainly testing the path validation logic path in SessionManager.
        # However, SessionManager tries to Document(path) if it exists.
        # An empty file is not a valid docx.
        # So we need to create a valid docx first or handle the RuntimeError.

        # Let's just create a valid empty docx first using the manager itself without path
        sid = manager.create_session()
        manager.get_session(sid).document.save(tmp_path)
        manager.close_session(sid)

        # Now load it back
        session_id = manager.create_session(file_path=tmp_path)
        assert session_id is not None
        manager.close_session(session_id)

    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
