"""Table row and column manipulation tools"""
import logging
from docx.table import Table
from docx_editor_skill.core.xml_util import ElementManipulator
from docx_editor_skill.core.registry_cleaner import RegistryCleaner
from docx_editor_skill.core.response import (
    create_context_aware_response,
    create_error_response,
)
from docx_editor_skill.utils.session_helpers import get_active_session

logger = logging.getLogger(__name__)


def docx_insert_row_at(
        table_id: str,
    position: str,
    row_index: int = None,
    copy_format: bool = False
) -> str:
    """
    Insert a new row at a specific position in a table.

    Supports precise row insertion with optional format copying from adjacent rows.

    Args:
        session_id (str): Active session ID returned by docx_create().
        table_id (str): Table element_id (e.g., "table_abc123").
        position (str): Insertion position string:
            - "after:N" - Insert after row N (0-based index)
            - "before:N" - Insert before row N
            - "start:table_id" - Insert at table start
            - "end:table_id" - Insert at table end
        row_index (int, optional): Direct row index (alternative to position).
        copy_format (bool): If True, copy formatting from adjacent row (default: False).

    Returns:
        str: JSON response with table_id, new_row_count, inserted_at, and cursor info.

    Example:
        >>> result = docx_insert_row_at(session_id, "table_123", "after:1", copy_format=True)
        >>> data = json.loads(result)
        >>> print(data["data"]["new_row_count"])  # 4 (if table had 3 rows)
    """


    session, error = get_active_session()
    if error:
        return error
    logger.debug(f"docx_insert_row_at called: session_id={session.session_id}, table_id={table_id}, position={position}, row_index={row_index}, copy_format={copy_format}")

    # Get the table object
    try:

        table = session.get_object(table_id)

    except ValueError as e:

        if "Special ID" in str(e) or "not available" in str(e):

            return create_error_response(str(e), error_type="SpecialIDNotAvailable")

        raise

    if not table:
        return create_error_response(f"Table {table_id} not found", error_type="ElementNotFound")

    if not isinstance(table, Table):
        return create_error_response(f"Element {table_id} is not a table", error_type="InvalidElementType")

    try:
        # Determine insertion index
        insert_index = None
        copy_format_from = None

        if row_index is not None:
            # Direct index provided
            insert_index = row_index
        else:
            # Parse position string
            if position.startswith("after:"):
                ref_index = int(position.split(":")[1])
                insert_index = ref_index + 1
                if copy_format:
                    copy_format_from = ref_index
            elif position.startswith("before:"):
                ref_index = int(position.split(":")[1])
                insert_index = ref_index
                if copy_format:
                    copy_format_from = ref_index
            elif position.startswith("start:"):
                insert_index = 0
                if copy_format and len(table.rows) > 0:
                    copy_format_from = 0
            elif position.startswith("end:"):
                insert_index = len(table.rows)
                if copy_format and len(table.rows) > 0:
                    copy_format_from = len(table.rows) - 1
            else:
                return create_error_response(
                    f"Invalid position format: {position}. Expected 'after:N', 'before:N', 'start:table_id', or 'end:table_id'",
                    error_type="ValidationError"
                )

        # Insert the row
        ElementManipulator.insert_row_at(table, insert_index, copy_format_from=copy_format_from)

        # Update context
        session.update_context(table_id, action="modify")
        session.cursor.element_id = table_id
        session.cursor.position = "inside_end"

        # Build response
        return create_context_aware_response(
            session,
            message=f"Row inserted at position {insert_index}",
            element_id=table_id,
            table_id=table_id,
            new_row_count=len(table.rows),
            inserted_at=insert_index,
            copy_format=copy_format
        )

    except IndexError as e:
        logger.exception(f"Index error in docx_insert_row_at: {e}")
        return create_error_response(
            f"Row index out of range: {str(e)}",
            error_type="IndexError"
        )
    except ValueError as e:
        logger.exception(f"Value error in docx_insert_row_at: {e}")
        return create_error_response(
            f"Invalid position: {str(e)}",
            error_type="ValidationError"
        )
    except Exception as e:
        logger.exception(f"docx_insert_row_at failed: {e}")
        return create_error_response(
            f"Failed to insert row: {str(e)}",
            error_type="CreationError"
        )


def docx_insert_col_at(
        table_id: str,
    position: str,
    col_index: int = None,
    copy_format: bool = False
) -> str:
    """
    Insert a new column at a specific position in a table.

    Supports precise column insertion with optional format copying from adjacent columns.

    Args:
        session_id (str): Active session ID returned by docx_create().
        table_id (str): Table element_id (e.g., "table_abc123").
        position (str): Insertion position string:
            - "after:N" - Insert after column N (0-based index)
            - "before:N" - Insert before column N
            - "start:table_id" - Insert at table start
            - "end:table_id" - Insert at table end
        col_index (int, optional): Direct column index (alternative to position).
        copy_format (bool): If True, copy formatting from adjacent column (default: False).

    Returns:
        str: JSON response with table_id, new_col_count, inserted_at, and cursor info.

    Example:
        >>> result = docx_insert_col_at(session_id, "table_123", "before:1", copy_format=True)
        >>> data = json.loads(result)
        >>> print(data["data"]["new_col_count"])  # 4 (if table had 3 columns)
    """


    session, error = get_active_session()
    if error:
        return error
    logger.debug(f"docx_insert_col_at called: session_id={session.session_id}, table_id={table_id}, position={position}, col_index={col_index}, copy_format={copy_format}")

    # Get the table object
    try:

        table = session.get_object(table_id)

    except ValueError as e:

        if "Special ID" in str(e) or "not available" in str(e):

            return create_error_response(str(e), error_type="SpecialIDNotAvailable")

        raise

    if not table:
        return create_error_response(f"Table {table_id} not found", error_type="ElementNotFound")

    if not isinstance(table, Table):
        return create_error_response(f"Element {table_id} is not a table", error_type="InvalidElementType")

    try:
        # Determine insertion index
        insert_index = None
        copy_format_from = None

        if col_index is not None:
            # Direct index provided
            insert_index = col_index
        else:
            # Parse position string
            if position.startswith("after:"):
                ref_index = int(position.split(":")[1])
                insert_index = ref_index + 1
                if copy_format:
                    copy_format_from = ref_index
            elif position.startswith("before:"):
                ref_index = int(position.split(":")[1])
                insert_index = ref_index
                if copy_format:
                    copy_format_from = ref_index
            elif position.startswith("start:"):
                insert_index = 0
                if copy_format and len(table.columns) > 0:
                    copy_format_from = 0
            elif position.startswith("end:"):
                insert_index = len(table.columns)
                if copy_format and len(table.columns) > 0:
                    copy_format_from = len(table.columns) - 1
            else:
                return create_error_response(
                    f"Invalid position format: {position}. Expected 'after:N', 'before:N', 'start:table_id', or 'end:table_id'",
                    error_type="ValidationError"
                )

        # Insert the column
        new_col_count = ElementManipulator.insert_col_at(table, insert_index, copy_format_from=copy_format_from)

        # Update context
        session.update_context(table_id, action="modify")
        session.cursor.element_id = table_id
        session.cursor.position = "inside_end"

        # Build response
        return create_context_aware_response(
            session,
            message=f"Column inserted at position {insert_index}",
            element_id=table_id,
            table_id=table_id,
            new_col_count=new_col_count,
            inserted_at=insert_index,
            copy_format=copy_format
        )

    except IndexError as e:
        logger.exception(f"Index error in docx_insert_col_at: {e}")
        return create_error_response(
            f"Column index out of range: {str(e)}",
            error_type="IndexError"
        )
    except ValueError as e:
        logger.exception(f"Value error in docx_insert_col_at: {e}")
        return create_error_response(
            f"Invalid position: {str(e)}",
            error_type="ValidationError"
        )
    except Exception as e:
        logger.exception(f"docx_insert_col_at failed: {e}")
        return create_error_response(
            f"Failed to insert column: {str(e)}",
            error_type="CreationError"
        )


def docx_delete_row(
        table_id: str,
    row_index: int = None,
    row_id: str = None
) -> str:
    """
    Delete a row from a table.

    Supports deletion by index or element_id. Automatically cleans up invalidated
    element_id mappings in the session registry.

    Args:
        session_id (str): Active session ID returned by docx_create().
        table_id (str): Table element_id (e.g., "table_abc123").
        row_index (int, optional): Row index to delete (0-based).
        row_id (str, optional): Row element_id to delete (alternative to row_index).

    Returns:
        str: JSON response with table_id, new_row_count, deleted_index,
             invalidated_ids, and cursor info.

    Raises:
        ValidationError: If attempting to delete the last row.
        IndexError: If row_index is out of range.
        ElementNotFound: If row_id is not found.

    Example:
        >>> result = docx_delete_row(session_id, "table_123", row_index=1)
        >>> data = json.loads(result)
        >>> print(data["data"]["invalidated_ids"])  # ["cell_xyz789", "cell_xyz790"]
    """


    session, error = get_active_session()
    if error:
        return error
    logger.debug(f"docx_delete_row called: session_id={session.session_id}, table_id={table_id}, row_index={row_index}, row_id={row_id}")

    # Get the table object
    try:

        table = session.get_object(table_id)

    except ValueError as e:

        if "Special ID" in str(e) or "not available" in str(e):

            return create_error_response(str(e), error_type="SpecialIDNotAvailable")

        raise

    if not table:
        return create_error_response(f"Table {table_id} not found", error_type="ElementNotFound")

    if not isinstance(table, Table):
        return create_error_response(f"Element {table_id} is not a table", error_type="InvalidElementType")

    try:
        # Determine deletion index
        delete_index = None

        if row_index is not None:
            delete_index = row_index
        elif row_id is not None:
            # Find the row by element_id (not commonly used, but supported)
            # Note: Rows typically don't have element_ids, but we support it for consistency
            return create_error_response(
                "Deletion by row_id is not supported (rows don't have element_ids). Use row_index instead.",
                error_type="ValidationError"
            )
        else:
            return create_error_response(
                "Either row_index or row_id must be provided",
                error_type="ValidationError"
            )

        # Find invalidated cell IDs before deletion
        invalidated_ids = RegistryCleaner.find_invalidated_ids(session, table, row_index=delete_index)

        # Delete the row
        ElementManipulator.delete_row(table, delete_index)

        # Clean up invalidated IDs
        RegistryCleaner.invalidate_ids(session, invalidated_ids)

        # Update context
        session.update_context(table_id, action="modify")
        session.cursor.element_id = table_id
        session.cursor.position = "inside_end"

        # Build response
        return create_context_aware_response(
            session,
            message=f"Row deleted at index {delete_index}",
            element_id=table_id,
            table_id=table_id,
            new_row_count=len(table.rows),
            deleted_index=delete_index,
            invalidated_ids=invalidated_ids
        )

    except ValueError as e:
        logger.exception(f"Validation error in docx_delete_row: {e}")
        return create_error_response(
            str(e),
            error_type="ValidationError"
        )
    except IndexError as e:
        logger.exception(f"Index error in docx_delete_row: {e}")
        return create_error_response(
            f"Row index out of range: {str(e)}",
            error_type="IndexError"
        )
    except Exception as e:
        logger.exception(f"docx_delete_row failed: {e}")
        return create_error_response(
            f"Failed to delete row: {str(e)}",
            error_type="UpdateError"
        )


def docx_delete_col(
        table_id: str,
    col_index: int = None,
    col_id: str = None
) -> str:
    """
    Delete a column from a table.

    Supports deletion by index or element_id. Automatically cleans up invalidated
    element_id mappings in the session registry.

    Args:
        session_id (str): Active session ID returned by docx_create().
        table_id (str): Table element_id (e.g., "table_abc123").
        col_index (int, optional): Column index to delete (0-based).
        col_id (str, optional): Column element_id to delete (alternative to col_index).

    Returns:
        str: JSON response with table_id, new_col_count, deleted_index,
             invalidated_ids, and cursor info.

    Raises:
        ValidationError: If attempting to delete the last column.
        IndexError: If col_index is out of range.
        ElementNotFound: If col_id is not found.

    Example:
        >>> result = docx_delete_col(session_id, "table_123", col_index=0)
        >>> data = json.loads(result)
        >>> print(data["data"]["new_col_count"])  # 2 (if table had 3 columns)
    """


    session, error = get_active_session()
    if error:
        return error
    logger.debug(f"docx_delete_col called: session_id={session.session_id}, table_id={table_id}, col_index={col_index}, col_id={col_id}")

    # Get the table object
    try:

        table = session.get_object(table_id)

    except ValueError as e:

        if "Special ID" in str(e) or "not available" in str(e):

            return create_error_response(str(e), error_type="SpecialIDNotAvailable")

        raise

    if not table:
        return create_error_response(f"Table {table_id} not found", error_type="ElementNotFound")

    if not isinstance(table, Table):
        return create_error_response(f"Element {table_id} is not a table", error_type="InvalidElementType")

    try:
        # Determine deletion index
        delete_index = None

        if col_index is not None:
            delete_index = col_index
        elif col_id is not None:
            # Find the column by element_id (not commonly used, but supported)
            # Note: Columns typically don't have element_ids, but we support it for consistency
            return create_error_response(
                "Deletion by col_id is not supported (columns don't have element_ids). Use col_index instead.",
                error_type="ValidationError"
            )
        else:
            return create_error_response(
                "Either col_index or col_id must be provided",
                error_type="ValidationError"
            )

        # Find invalidated cell IDs before deletion
        invalidated_ids = RegistryCleaner.find_invalidated_ids(session, table, col_index=delete_index)

        # Delete the column
        new_col_count = ElementManipulator.delete_col(table, delete_index)

        # Clean up invalidated IDs
        RegistryCleaner.invalidate_ids(session, invalidated_ids)

        # Update context
        session.update_context(table_id, action="modify")
        session.cursor.element_id = table_id
        session.cursor.position = "inside_end"

        # Build response
        return create_context_aware_response(
            session,
            message=f"Column deleted at index {delete_index}",
            element_id=table_id,
            table_id=table_id,
            new_col_count=new_col_count,
            deleted_index=delete_index,
            invalidated_ids=invalidated_ids
        )

    except ValueError as e:
        logger.exception(f"Validation error in docx_delete_col: {e}")
        return create_error_response(
            str(e),
            error_type="ValidationError"
        )
    except IndexError as e:
        logger.exception(f"Index error in docx_delete_col: {e}")
        return create_error_response(
            f"Column index out of range: {str(e)}",
            error_type="IndexError"
        )
    except Exception as e:
        logger.exception(f"docx_delete_col failed: {e}")
        return create_error_response(
            f"Failed to delete column: {str(e)}",
            error_type="UpdateError"
        )