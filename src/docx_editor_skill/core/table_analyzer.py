"""
Table structure analyzer for detecting irregular structures and generating visualizations.

This module provides utilities for analyzing table structures, detecting merged cells,
nested tables, and generating ASCII visualizations.
"""

import logging
from typing import Dict, Any, List, Tuple, Optional
from docx.table import Table

logger = logging.getLogger(__name__)


class TableStructureAnalyzer:
    """Analyzer for table structures."""

    @staticmethod
    def detect_irregular_structure(table: Table, session: Optional[Any] = None) -> Dict[str, Any]:
        """
        Detect if table contains irregular structures.

        Args:
            table: Table object to analyze
            session: Optional Session object for cell ID registration

        Returns:
            Dictionary with detection results and optionally grid_structure with cell_ids
        """
        result = {
            "is_irregular": False,
            "has_merged_cells": False,
            "has_nested_tables": False,
            "row_col_inconsistent": False,
            "irregular_regions": [],
            "grid_structure": []
        }

        if not table.rows:
            return result

        # Check for merged cells
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                # Check grid_span (column span)
                try:
                    tc_pr = cell._element.tcPr
                    if tc_pr is not None and tc_pr.gridSpan is not None:
                        grid_span = tc_pr.gridSpan.val
                        if grid_span > 1:
                            result["has_merged_cells"] = True
                            result["is_irregular"] = True
                            result["irregular_regions"].append({
                                "type": "merged",
                                "row": row_idx,
                                "col": col_idx,
                                "colspan": grid_span
                            })
                except Exception as e:
                    logger.debug(f"Error checking grid_span at ({row_idx}, {col_idx}): {e}")

                # Check vMerge (row span)
                try:
                    tc_pr = cell._element.tcPr
                    if tc_pr is not None and tc_pr.vMerge is not None:
                        result["has_merged_cells"] = True
                        result["is_irregular"] = True
                        result["irregular_regions"].append({
                            "type": "merged_vertical",
                            "row": row_idx,
                            "col": col_idx
                        })
                except Exception as e:
                    logger.debug(f"Error checking vMerge at ({row_idx}, {col_idx}): {e}")

        # Check for nested tables
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                if len(cell.tables) > 0:
                    result["has_nested_tables"] = True
                    result["is_irregular"] = True
                    result["irregular_regions"].append({
                        "type": "nested",
                        "row": row_idx,
                        "col": col_idx
                    })

        # Check row/column consistency
        col_counts = [len(row.cells) for row in table.rows]
        if len(set(col_counts)) > 1:
            result["row_col_inconsistent"] = True
            result["is_irregular"] = True

        # Generate grid_structure with cell_ids if session is provided
        if session is not None:
            grid_structure = []
            for row_idx, row in enumerate(table.rows):
                row_data = []
                for col_idx, cell in enumerate(row.cells):
                    cell_info = {
                        "row": row_idx,
                        "col": col_idx,
                        "text": cell.text[:50] if cell.text else ""
                    }

                    # Add cell_id if session is provided
                    try:
                        cell_id = session._get_element_id(cell, auto_register=True)
                        if cell_id:
                            cell_info["cell_id"] = cell_id
                    except Exception as e:
                        logger.warning(f"Failed to generate cell_id for cell ({row_idx}, {col_idx}): {e}")

                    row_data.append(cell_info)
                grid_structure.append(row_data)

            result["grid_structure"] = grid_structure

        return result

    @staticmethod
    def generate_ascii_visualization(table: Table) -> str:
        """
        Generate ASCII visualization of table structure.

        Args:
            table: Table object to visualize

        Returns:
            ASCII string representation
        """
        if not table.rows:
            return "Empty table (0 rows)"

        rows_count = len(table.rows)
        cols_count = len(table.columns) if table.rows else 0

        # Detect irregular structure
        structure_info = TableStructureAnalyzer.detect_irregular_structure(table)

        lines = []
        lines.append(f"Table: {rows_count} rows x {cols_count} cols "
                     f"(merged cells: {'yes' if structure_info['has_merged_cells'] else 'no'})")
        lines.append("")

        # Calculate column widths (fixed at 20 for simplicity)
        col_widths = [20] * cols_count

        # Generate table rows
        for row_idx, row in enumerate(table.rows):
            # Top border
            border = "+" + "+".join(["-" * (w + 2) for w in col_widths]) + "+"
            lines.append(border)

            # Cell contents
            cell_texts = []
            for col_idx, cell in enumerate(row.cells):
                text = cell.text.replace('\n', ' ')
                if len(text) > 20:
                    text = text[:17] + "..."
                if not text:
                    text = "[empty]"
                cell_texts.append(f" {text:<{col_widths[col_idx]}} ")

            lines.append("|" + "|".join(cell_texts) + "|")

        # Bottom border
        border = "+" + "+".join(["-" * (w + 2) for w in col_widths]) + "+"
        lines.append(border)

        return "\n".join(lines)

    @staticmethod
    def get_fillable_cells(table: Table, structure_info: Dict[str, Any] = None) -> List[Tuple[int, int]]:
        """
        Get list of fillable cell coordinates.

        Args:
            table: Table object
            structure_info: Optional pre-computed structure info

        Returns:
            List of (row, col) tuples for fillable cells
        """
        if not table.rows:
            return []

        if structure_info is None:
            structure_info = TableStructureAnalyzer.detect_irregular_structure(table)

        fillable_cells = []

        # If no irregular structure, all cells are fillable
        if not structure_info["is_irregular"]:
            for row_idx, row in enumerate(table.rows):
                for col_idx in range(len(row.cells)):
                    fillable_cells.append((row_idx, col_idx))
            return fillable_cells

        # Build set of irregular cell positions
        irregular_positions = set()
        for region in structure_info["irregular_regions"]:
            irregular_positions.add((region["row"], region["col"]))

        # Add all non-irregular cells
        for row_idx, row in enumerate(table.rows):
            for col_idx in range(len(row.cells)):
                if (row_idx, col_idx) not in irregular_positions:
                    fillable_cells.append((row_idx, col_idx))

        return fillable_cells
