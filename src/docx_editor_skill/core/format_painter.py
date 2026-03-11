from typing import Any, Union
import logging
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docx.table import Table

logger = logging.getLogger(__name__)

class FormatPainter:
    def copy_format(self, source: Union[Run, Paragraph, Table], target: Union[Run, Paragraph, Table]) -> None:
        """
        Copy format from source to target.
        Dispatches to specific methods based on type.
        """
        # Exact match dispatch
        if isinstance(source, Run) and isinstance(target, Run):
            self._copy_run_format(source, target)
        elif isinstance(source, Paragraph) and isinstance(target, Paragraph):
            self._copy_paragraph_format(source, target)
        elif isinstance(source, Table) and isinstance(target, Table):
            self._copy_table_format(source, target)

        # Smart matching (Cross-type)
        elif isinstance(source, Paragraph) and isinstance(target, Run):
            # Apply Paragraph Style's font to Run
            if source.style and hasattr(source.style, 'font'):
                self._copy_font_properties(source.style.font, target.font)
        elif isinstance(source, Run) and isinstance(target, Paragraph):
            # Apply Run's font to all runs in Paragraph
            for run in target.runs:
                self._copy_font_properties(source.font, run.font)

        else:
            raise ValueError(f"Unsupported object type or mismatch: {type(source)} -> {type(target)}")

    def _copy_font_properties(self, src_font: Any, tgt_font: Any) -> None:
        """Helper to copy font properties from one font object to another.

        Args:
            src_font: Source font object (typically Run.font)
            tgt_font: Target font object to modify
        """
        logger.debug(f"Copying font properties from {src_font.name} to target")

        # Font name and size
        tgt_font.name = src_font.name
        tgt_font.size = src_font.size

        # Boolean properties - basic
        tgt_font.bold = src_font.bold
        tgt_font.italic = src_font.italic
        tgt_font.underline = src_font.underline
        tgt_font.strike = src_font.strike
        tgt_font.superscript = src_font.superscript
        tgt_font.subscript = src_font.subscript

        # Boolean properties - advanced
        if hasattr(src_font, 'double_strike'):
            tgt_font.double_strike = src_font.double_strike
        if hasattr(src_font, 'all_caps'):
            tgt_font.all_caps = src_font.all_caps
        if hasattr(src_font, 'small_caps'):
            tgt_font.small_caps = src_font.small_caps
        if hasattr(src_font, 'shadow'):
            tgt_font.shadow = src_font.shadow
        if hasattr(src_font, 'outline'):
            tgt_font.outline = src_font.outline
        if hasattr(src_font, 'rtl'):
            tgt_font.rtl = src_font.rtl
        if hasattr(src_font, 'emboss'):
            tgt_font.emboss = src_font.emboss
        if hasattr(src_font, 'imprint'):
            tgt_font.imprint = src_font.imprint

        # Color
        if src_font.color and src_font.color.rgb:
            try:
                tgt_font.color.rgb = src_font.color.rgb
            except Exception as e:
                logger.warning(f"Failed to copy font color: {e}")

        # Highlight
        if hasattr(src_font, 'highlight_color'):
            try:
                tgt_font.highlight_color = src_font.highlight_color
            except Exception as e:
                logger.warning(f"Failed to copy highlight color: {e}")

    def _copy_run_format(self, source: Run, target: Run) -> None:
        """Copy all font properties from source run to target run."""
        self._copy_font_properties(source.font, target.font)

    def _copy_paragraph_format(self, source: Paragraph, target: Paragraph) -> None:
        """Copy paragraph formatting and style."""
        logger.debug("Copying paragraph format")
        src_fmt = source.paragraph_format
        tgt_fmt = target.paragraph_format

        try:
            # Alignment
            tgt_fmt.alignment = src_fmt.alignment

            # Indentation
            tgt_fmt.left_indent = src_fmt.left_indent
            tgt_fmt.right_indent = src_fmt.right_indent
            tgt_fmt.first_line_indent = src_fmt.first_line_indent

            # Spacing
            tgt_fmt.space_before = src_fmt.space_before
            tgt_fmt.space_after = src_fmt.space_after
            tgt_fmt.line_spacing = src_fmt.line_spacing
            tgt_fmt.line_spacing_rule = src_fmt.line_spacing_rule

            # Pagination
            tgt_fmt.keep_together = src_fmt.keep_together
            tgt_fmt.keep_with_next = src_fmt.keep_with_next
            tgt_fmt.page_break_before = src_fmt.page_break_before
            tgt_fmt.widow_control = src_fmt.widow_control

            # Tab stops (if available)
            if hasattr(src_fmt, 'tab_stops') and src_fmt.tab_stops:
                try:
                    tgt_fmt.tab_stops = src_fmt.tab_stops
                except Exception as e:
                    logger.warning(f"Failed to copy tab stops: {e}")
        except Exception as e:
            logger.error(f"Error copying paragraph format properties: {e}")
            raise

        # Style
        # Note: assigning style by name usually works in python-docx if style exists
        # We try to copy the style object or name
        try:
            target.style = source.style
            logger.debug(f"Copied style: {source.style}")
        except Exception as e:
            logger.warning(f"Failed to copy style '{source.style}': {e}")
            # Continue without failing - other properties were copied successfully

    def _copy_table_format(self, source: Table, target: Table) -> None:
        """
        Copy table formatting including style, borders, and shading.
        Uses low-level OXML manipulation for borders/shading as python-docx API is limited here.
        """
        import copy
        from docx.oxml.ns import qn

        logger.debug("Copying table format")

        try:
            # Copy Style
            target.style = source.style
        except Exception as e:
            logger.warning(f"Failed to copy table style: {e}")

        try:
            # OXML Manipulation for Borders and Shading
            # We act on tblPr (Table Properties)
            src_tblPr = source._element.tblPr
            tgt_tblPr = target._element.tblPr

            # 1. Borders (tblBorders)
            src_borders = src_tblPr.find(qn('w:tblBorders'))
            if src_borders is not None:
                try:
                    # Remove existing borders in target if any
                    old_borders = tgt_tblPr.find(qn('w:tblBorders'))
                    if old_borders is not None:
                        tgt_tblPr.remove(old_borders)

                    # Deep copy source borders and append to target
                    new_borders = copy.deepcopy(src_borders)
                    tgt_tblPr.append(new_borders)
                    logger.debug("Copied table borders")
                except Exception as e:
                    logger.warning(f"Failed to copy table borders: {e}")

            # 2. Shading (shd)
            src_shd = src_tblPr.find(qn('w:shd'))
            if src_shd is not None:
                try:
                    # Remove existing shading
                    old_shd = tgt_tblPr.find(qn('w:shd'))
                    if old_shd is not None:
                        tgt_tblPr.remove(old_shd)

                    # Deep copy
                    new_shd = copy.deepcopy(src_shd)
                    tgt_tblPr.append(new_shd)
                    logger.debug("Copied table shading")
                except Exception as e:
                    logger.warning(f"Failed to copy table shading: {e}")

            # Grid/Column widths (tblGrid) - Optional but good for layout consistency
            src_grid = source._element.find(qn('w:tblGrid'))
            if src_grid is not None:
                try:
                    old_grid = target._element.find(qn('w:tblGrid'))
                    if old_grid is not None:
                        target._element.remove(old_grid)
                    target._element.insert(0, copy.deepcopy(src_grid)) # tblGrid is usually early in sequence
                    logger.debug("Copied table grid")
                except Exception as e:
                    logger.warning(f"Failed to copy table grid: {e}")
        except Exception as e:
            logger.error(f"Error accessing table OXML structure: {e}")
            raise

    def copy_row_format(self, source_row, target_row) -> None:
        """
        Copy formatting from source row to target row (cell by cell).

        Args:
            source_row: Source _Row object
            target_row: Target _Row object
        """
        logger.debug("Copying row format")

        # Ensure both rows have the same number of cells
        min_cells = min(len(source_row.cells), len(target_row.cells))

        for i in range(min_cells):
            try:
                self.copy_cell_format(source_row.cells[i], target_row.cells[i])
            except Exception as e:
                logger.warning(f"Failed to copy format for cell {i}: {e}")

    def copy_col_format(self, table: Table, source_col_index: int, target_col_index: int) -> None:
        """
        Copy formatting from source column to target column (cell by cell).

        Args:
            table: Table object
            source_col_index: Source column index (0-based)
            target_col_index: Target column index (0-based)
        """
        logger.debug(f"Copying column format from {source_col_index} to {target_col_index}")

        for row in table.rows:
            try:
                if source_col_index < len(row.cells) and target_col_index < len(row.cells):
                    self.copy_cell_format(row.cells[source_col_index], row.cells[target_col_index])
            except Exception as e:
                logger.warning(f"Failed to copy format for row cell: {e}")

    def copy_cell_format(self, source_cell, target_cell) -> None:
        """
        Copy formatting from source cell to target cell.

        Includes:
        - Paragraph formatting (alignment, spacing)
        - Run formatting (font, size, color, bold, italic)
        - Cell background color
        - Cell borders

        Args:
            source_cell: Source _Cell object
            target_cell: Target _Cell object
        """
        import copy
        from docx.oxml.ns import qn

        logger.debug("Copying cell format")

        try:
            # Copy paragraph formatting from first paragraph if exists
            if len(source_cell.paragraphs) > 0 and len(target_cell.paragraphs) > 0:
                src_para = source_cell.paragraphs[0]
                tgt_para = target_cell.paragraphs[0]
                self._copy_paragraph_format(src_para, tgt_para)

                # Copy run formatting if both have runs
                if len(src_para.runs) > 0 and len(tgt_para.runs) > 0:
                    self._copy_run_format(src_para.runs[0], tgt_para.runs[0])

            # Copy cell properties (background color, borders)
            src_tcPr = source_cell._tc.get_or_add_tcPr()
            tgt_tcPr = target_cell._tc.get_or_add_tcPr()

            # Copy shading (background color)
            src_shd = src_tcPr.find(qn('w:shd'))
            if src_shd is not None:
                try:
                    old_shd = tgt_tcPr.find(qn('w:shd'))
                    if old_shd is not None:
                        tgt_tcPr.remove(old_shd)
                    tgt_tcPr.append(copy.deepcopy(src_shd))
                    logger.debug("Copied cell shading")
                except Exception as e:
                    logger.warning(f"Failed to copy cell shading: {e}")

            # Copy borders
            src_borders = src_tcPr.find(qn('w:tcBorders'))
            if src_borders is not None:
                try:
                    old_borders = tgt_tcPr.find(qn('w:tcBorders'))
                    if old_borders is not None:
                        tgt_tcPr.remove(old_borders)
                    tgt_tcPr.append(copy.deepcopy(src_borders))
                    logger.debug("Copied cell borders")
                except Exception as e:
                    logger.warning(f"Failed to copy cell borders: {e}")

        except Exception as e:
            logger.error(f"Error copying cell format: {e}")
            raise

