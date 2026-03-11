import platform
import re
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

def validate_path_safety(file_path: str) -> None:
    """
    Validate that the path is safe and appropriate for the current OS.
    Raises ValueError if the path looks suspicious (e.g., Windows path on Linux).
    """
    system = platform.system()

    # 1. Check for Windows paths on Linux/macOS
    if system != "Windows":
        # Check for backslashes (strong indicator of Windows path)
        if "\\" in file_path:
            logger.error(f"Path validation failed: Windows path on {system}: {file_path}")
            raise ValueError(
                f"Path '{file_path}' contains backslashes. "
                "This suggests a Windows path, but the server is running on "
                f"{system}. Please use forward slashes."
            )

        # Check for drive letters (e.g. C:/Users)
        # Regex: Start of string, single letter, colon, slash
        if re.match(r'^[a-zA-Z]:/', file_path):
            logger.error(f"Path validation failed: Drive letter on {system}: {file_path}")
            raise ValueError(
                f"Path '{file_path}' starts with a drive letter. "
                "This suggests a Windows path, but the server is running on "
                f"{system}."
            )

    # 2. Check for Linux absolute paths on Windows
    if system == "Windows":
        # Check for paths starting with / (Posix absolute) without a drive
        # While technically valid (root-relative) on Windows, usually implies a mistake
        # when coming from a cross-platform client context.
        # pathlib.Path('/foo').is_absolute() is False on Windows.
        p = Path(file_path)
        if file_path.startswith("/") and not p.is_absolute():
            # It starts with / but isn't a full absolute path (no drive/UNC)
            logger.error(f"Path validation failed: Linux path on Windows: {file_path}")
            raise ValueError(
                f"Path '{file_path}' looks like a Linux/Unix absolute path (missing drive letter). "
                "The server is running on Windows. "
                "Please provide a full Windows path (e.g., 'C:\\Users\\...')."
            )
