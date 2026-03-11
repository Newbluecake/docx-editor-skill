"""Registry cleaner for managing element_id mappings after deletions."""
import logging
from typing import List, Optional, TYPE_CHECKING
from docx.table import Table

if TYPE_CHECKING:
    from docx_editor_skill.core.session import Session

logger = logging.getLogger(__name__)


class RegistryCleaner:
    """
    Handles cleanup of element_id mappings when table rows/columns are deleted.
    """

    @staticmethod
    def find_invalidated_ids(
        session: 'Session',
        table: Table,
        row_index: Optional[int] = None,
        col_index: Optional[int] = None
    ) -> List[str]:
        """
        Find element_ids that will become invalid after deleting a row or column.

        Args:
            session: Session object containing the object_registry
            table: Table object being modified
            row_index: Optional row index being deleted
            col_index: Optional column index being deleted

        Returns:
            List of element_ids that will be invalidated
        """
        invalidated_ids = []

        # Iterate through the object registry to find cell IDs
        for element_id, obj in list(session.object_registry.items()):
            # Check if this is a cell ID
            if not element_id.startswith('cell_'):
                continue

            # Try to determine if this cell belongs to the deleted row/column
            try:
                # Check if the object is still a valid cell
                if not hasattr(obj, '_tc'):
                    continue

                # Find which table this cell belongs to
                cell_table = None
                for row in table.rows:
                    if obj in row.cells:
                        cell_table = table
                        break

                if cell_table != table:
                    # This cell doesn't belong to the table being modified
                    continue

                # Find the cell's position in the table
                for r_idx, row in enumerate(table.rows):
                    for c_idx, cell in enumerate(row.cells):
                        if cell == obj:
                            # Check if this cell is in the deleted row or column
                            if row_index is not None and r_idx == row_index:
                                invalidated_ids.append(element_id)
                                logger.debug(f"Cell {element_id} at row {r_idx} will be invalidated")
                            elif col_index is not None and c_idx == col_index:
                                invalidated_ids.append(element_id)
                                logger.debug(f"Cell {element_id} at col {c_idx} will be invalidated")
                            break

            except Exception as e:
                logger.warning(f"Error checking cell {element_id}: {e}")
                continue

        return invalidated_ids

    @staticmethod
    def invalidate_ids(session: 'Session', ids: List[str]) -> None:
        """
        Remove element_ids from the object registry.

        Args:
            session: Session object containing the object_registry
            ids: List of element_ids to remove
        """
        for element_id in ids:
            if element_id in session.object_registry:
                del session.object_registry[element_id]
                logger.debug(f"Removed {element_id} from object registry")
