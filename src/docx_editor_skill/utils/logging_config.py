import logging
import traceback
from pathlib import Path
from logging.handlers import RotatingFileHandler


class StackTraceFormatter(logging.Formatter):
    """Custom formatter that adds call stack to ERROR level logs."""

    def format(self, record):
        result = super().format(record)

        # Add call stack for ERROR level logs without exception info
        if record.levelno >= logging.ERROR and not record.exc_info:
            stack = ''.join(traceback.format_stack()[:-2])  # Exclude formatter frames
            result += f'\nCall Stack:\n{stack}'

        return result


def setup_file_logging(
    log_dir: str = "./logs",
    max_bytes: int = 10485760,  # 10MB
    backup_count: int = 5,
    log_level: int = logging.INFO
) -> bool:
    """
    Setup file logging with rotation for the root logger.

    Args:
        log_dir: Directory to store log files (default: ./logs)
        max_bytes: Maximum size of each log file in bytes (default: 10MB)
        backup_count: Number of backup files to keep (default: 5)
        log_level: Logging level (default: INFO)

    Returns:
        bool: True if file logging was successfully configured, False otherwise
    """
    try:
        # Create log directory if it doesn't exist
        log_path = Path(log_dir)
        log_path.mkdir(parents=True, exist_ok=True)

        # Define log file path
        log_file = log_path / "docx-editor-skill.log"

        # Create rotating file handler
        file_handler = RotatingFileHandler(
            log_file,
            maxBytes=max_bytes,
            backupCount=backup_count,
            encoding='utf-8'
        )

        # Set formatter
        formatter = StackTraceFormatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        file_handler.setFormatter(formatter)
        file_handler.setLevel(log_level)

        # Add handler to root logger
        root_logger = logging.getLogger()
        root_logger.addHandler(file_handler)

        return True

    except PermissionError:
        # Log to console if file logging fails due to permissions
        logging.error(f"Permission denied: Cannot write to log directory '{log_dir}'. File logging disabled.")
        return False
    except Exception as e:
        # Log any other errors to console
        logging.error(f"Failed to setup file logging: {e}. File logging disabled.")
        return False
