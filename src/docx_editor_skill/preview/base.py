from abc import ABC, abstractmethod
import logging

logger = logging.getLogger(__name__)

class PreviewController(ABC):
    """Abstract base class for controlling the Word application preview."""

    @abstractmethod
    def prepare_for_save(self, file_path: str) -> bool:
        """
        Prepare the document for saving.
        This usually involves closing the document in the editor to release the file lock.

        Args:
            file_path: The absolute path to the file about to be saved.

        Returns:
            True if the document was open and successfully closed (action taken),
            False otherwise (no action taken).
        """
        pass

    @abstractmethod
    def refresh(self, file_path: str) -> bool:
        """
        Refresh the preview by reloading the document in the editor.

        Args:
            file_path: The absolute path to the file that was just saved.

        Returns:
            True if successful, False otherwise.
        """
        pass

class NoOpPreviewController(PreviewController):
    """Null implementation for platforms where preview is not supported (Linux/macOS) or disabled."""

    def prepare_for_save(self, file_path: str) -> bool:
        # No-op
        return False

    def refresh(self, file_path: str) -> bool:
        # No-op
        return False
