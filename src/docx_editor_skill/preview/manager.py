import sys
import logging
from typing import Optional
from .base import PreviewController, NoOpPreviewController

logger = logging.getLogger(__name__)

class PreviewManager:
    """Factory for creating the appropriate PreviewController."""

    _instance: Optional[PreviewController] = None

    @staticmethod
    def get_controller() -> PreviewController:
        """
        Get the singleton instance of the PreviewController.
        Returns a Win32PreviewController on Windows if available,
        otherwise returns a NoOpPreviewController.
        """
        if PreviewManager._instance is None:
            if sys.platform == "win32":
                try:
                    # Dynamic import to avoid loading win32 dependencies on non-Windows
                    # or if the module hasn't been implemented yet (T-003)
                    from .win32 import Win32PreviewController
                    PreviewManager._instance = Win32PreviewController()
                    logger.info("Initialized Win32PreviewController for live preview")
                except ImportError:
                    logger.debug("Win32PreviewController module not found. Falling back to NoOp.")
                    PreviewManager._instance = NoOpPreviewController()
                except Exception as e:
                    logger.warning(f"Failed to initialize Win32PreviewController: {e}. Falling back to NoOp.")
                    PreviewManager._instance = NoOpPreviewController()
            else:
                logger.debug(f"Platform '{sys.platform}' does not support live preview. Using NoOp.")
                PreviewManager._instance = NoOpPreviewController()

        return PreviewManager._instance

    @staticmethod
    def set_controller(controller: PreviewController):
        """Allow injecting a controller (useful for testing)."""
        PreviewManager._instance = controller
