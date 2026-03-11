import logging
import sys
import io
from docx_editor_skill.utils.logging_config import StackTraceFormatter


LEVEL_NAMES = {
    "CRITICAL": logging.CRITICAL,
    "ERROR": logging.ERROR,
    "WARNING": logging.WARNING,
    "INFO": logging.INFO,
    "DEBUG": logging.DEBUG,
}


def parse_log_level(level: str) -> int:
    """Convert a log level string to logging constant, raising ValueError on invalid input."""
    if isinstance(level, str):
        upper = level.strip().upper()
        if upper in LEVEL_NAMES:
            return LEVEL_NAMES[upper]
    raise ValueError(f"Invalid log level: {level}")


def set_global_log_level(level: str | int) -> int:
    """Set root and docx-editor-skill logger levels; returns normalized numeric level."""
    numeric_level = level if isinstance(level, int) else parse_log_level(level)

    root_logger = logging.getLogger()
    root_logger.setLevel(numeric_level)

    server_logger = logging.getLogger("docx-editor-skill")
    server_logger.setLevel(numeric_level)

    return numeric_level


def get_global_log_level() -> str:
    """Return the current effective root log level name (e.g., INFO)."""
    level = logging.getLogger().getEffectiveLevel()
    for name, value in LEVEL_NAMES.items():
        if level == value:
            return name
    return logging.getLevelName(level)


def setup_logger(name: str, level: int = logging.INFO) -> logging.Logger:
    """
    Setup a logger that writes to stderr (to avoid interfering with MCP stdout communication).
    """
    logger = logging.getLogger(name)
    logger.setLevel(level)

    if not logger.handlers:
        # MCP servers communicate via stdout, so logs MUST go to stderr
        stream = sys.stderr
        try:
            stream = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")  # type: ignore[arg-type]
        except Exception:
            stream = sys.stderr
        handler = logging.StreamHandler(stream)
        formatter = StackTraceFormatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)

    return logger
