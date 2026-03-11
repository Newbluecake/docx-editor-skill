"""Unit tests for stacktrace logging functionality"""
import logging
import pytest
from io import StringIO
from docx_editor_skill.utils.logging_config import StackTraceFormatter
from docx_editor_skill.utils.logger import setup_logger


def test_stacktrace_formatter_adds_stack_to_error():
    """Test that StackTraceFormatter adds call stack to ERROR logs without exception info"""
    formatter = StackTraceFormatter('%(levelname)s - %(message)s')
    record = logging.LogRecord(
        name='test', level=logging.ERROR, pathname='test.py', lineno=1,
        msg='Test error', args=(), exc_info=None
    )

    result = formatter.format(record)

    assert 'Test error' in result
    assert 'Call Stack:' in result
    # Verify stack trace contains pytest or python frames
    assert 'pytest' in result or 'python' in result.lower()


def test_stacktrace_formatter_skips_stack_for_exception():
    """Test that StackTraceFormatter does not add call stack when exception info is present"""
    formatter = StackTraceFormatter('%(levelname)s - %(message)s')

    try:
        raise ValueError("Test exception")
    except ValueError:
        import sys
        record = logging.LogRecord(
            name='test', level=logging.ERROR, pathname='test.py', lineno=1,
            msg='Test error', args=(), exc_info=sys.exc_info()
        )

    result = formatter.format(record)

    assert 'Test error' in result
    assert 'Call Stack:' not in result  # Should not add call stack when exc_info is present


def test_stacktrace_formatter_skips_stack_for_info():
    """Test that StackTraceFormatter does not add call stack to INFO logs"""
    formatter = StackTraceFormatter('%(levelname)s - %(message)s')
    record = logging.LogRecord(
        name='test', level=logging.INFO, pathname='test.py', lineno=1,
        msg='Test info', args=(), exc_info=None
    )

    result = formatter.format(record)

    assert 'Test info' in result
    assert 'Call Stack:' not in result


def test_setup_logger_uses_stacktrace_formatter():
    """Test that setup_logger configures StackTraceFormatter"""
    logger = setup_logger('test_logger', level=logging.ERROR)

    # Check that the logger has a handler with StackTraceFormatter
    assert len(logger.handlers) > 0
    handler = logger.handlers[0]
    assert isinstance(handler.formatter, StackTraceFormatter)


def test_logger_exception_includes_traceback():
    """Test that logger.exception() includes traceback information"""
    stream = StringIO()
    handler = logging.StreamHandler(stream)
    handler.setFormatter(StackTraceFormatter('%(levelname)s - %(message)s'))

    logger = logging.getLogger('test_exception')
    logger.setLevel(logging.ERROR)
    logger.handlers = [handler]

    try:
        raise ValueError("Test exception")
    except ValueError:
        logger.exception("Caught exception")

    output = stream.getvalue()

    assert 'Caught exception' in output
    assert 'ValueError: Test exception' in output
    assert 'Traceback' in output
