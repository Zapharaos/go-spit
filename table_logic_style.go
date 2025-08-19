// table_logic_style.go - Table cell styling logic.
//
// This file implements algorithms and operations for applying styles, borders, and formatting to exported tables.
// Styling errors are logged as warnings, allowing export to continue with best-effort visual enhancements.
// Used by spreadsheet backends to support advanced styling features in exported files.

package spit

import (
	"fmt"
)

// renderStyles applies all styling and border operations to the table.
// It processes header styles, data cell styles, column borders, row borders, and cell-specific borders in order.
// Errors are wrapped and returned, but processing continues for best-effort styling.
func (t *Table) renderStyles(ops TableOperations) error {
	dataStartRow := t.getDataStartRow()
	totalColumns := t.Columns.getTotalColumnCount()
	dataEndRow := dataStartRow + len(t.Data) - 1

	// Apply header styles and borders
	if t.WriteHeader && len(t.Columns) > 0 {
		if err := t.applyHeaderStyles(ops); err != nil {
			return fmt.Errorf("failed to apply header styles: %w", err)
		}
	}

	// Apply data cell styles
	if err := t.applyCellStyles(dataStartRow, dataEndRow, ops); err != nil {
		return fmt.Errorf("failed to apply cell styles: %w", err)
	}

	// Apply column borders
	if err := t.applyColumnBorders(dataStartRow, dataEndRow, ops); err != nil {
		return fmt.Errorf("failed to apply column borders: %w", err)
	}

	// Apply row borders for each data row
	for rowIndex := range t.Data {
		actualRowNum := rowIndex + dataStartRow
		err := t.applyRowBorders(rowIndex, actualRowNum, totalColumns, ops)
		if err != nil {
			return fmt.Errorf("failed to apply row borders: %w", err)
		}
	}

	// Apply cell-specific borders last to override other border settings
	if err := t.applyCellSpecificBorders(dataStartRow, ops); err != nil {
		return fmt.Errorf("failed to apply cell-specific borders: %w", err)
	}

	return nil
}

// applyHeaderStyles applies styling and borders to header rows
func (t *Table) applyHeaderStyles(ops TableOperations) error {
	maxDepth := t.Columns.getMaxDepth()
	totalColumns := t.Columns.getTotalColumnCount()

	borders := NewBorderOptions(BorderStyleThin)

	// Apply bottom border to each header row
	for row := 1; row <= maxDepth; row++ {
		for col := 1; col <= totalColumns; col++ {
			if err := t.applyBordersToCell(col, row, borders, ops); err != nil {
				L().Warn("Failed to apply header cell-specific border",
					Int("column", col),
					Int("row", row),
					Error(err))
			}
		}
	}

	// Apply header cell styles
	if err := t.applyHeaderCellStyles(ops); err != nil {
		return fmt.Errorf("failed to apply header cell styles: %w", err)
	}

	return nil
}

// applyHeaderCellStyles applies default styling to all header cells.
// Styles include bold text, background color, and centered alignment.
func (t *Table) applyHeaderCellStyles(ops TableOperations) error {
	maxDepth := t.Columns.getMaxDepth()
	totalColumns := t.Columns.getTotalColumnCount()

	// Default header style configuration
	headerStyle := Style{
		Bold:            true,
		BackgroundColor: "#E0E0E0",
		Alignment:       AlignmentCenterMiddle,
	}

	// Apply header styling to all header rows
	if err := ops.applyStyleToRange(1, 1, totalColumns, maxDepth, headerStyle); err != nil {
		L().Warn("Failed to apply header range style", Error(err))
		return err
	}

	return nil
}

// applyCellStyles applies styling to all data cells based on priority: cell > row > column.
// For each cell, determines the most specific style to apply and applies it.
func (t *Table) applyCellStyles(dataStartRow, dataEndRow int, ops TableOperations) error {
	flatColumns := t.Columns.getFlattenedColumns()

	// Apply styles to each data row
	for rowIndex := dataStartRow; rowIndex <= dataEndRow; rowIndex++ {
		dataRowIndex := t.getDataIndexFromRowIndex(rowIndex)

		// Skip if data row index is out of bounds
		if dataRowIndex >= len(t.Data) {
			break
		}

		// Get row-level style if configured
		var rowStyle *Style
		if rc, exists := t.RowOptionsMap[dataRowIndex]; exists && rc.Style != nil {
			rowStyle = rc.Style
		}

		// Process each column in this row
		for colIndex, column := range flatColumns {
			actualColIndex := colIndex + 1
			var styleToApply *Style

			// Style priority: Cell-specific > Row-specific > Column-specific
			if cc, exists := t.CellOptionsMap[actualColIndex]; exists {
				if cellOptions, cellExists := cc[dataRowIndex]; cellExists && cellOptions.Style != nil {
					styleToApply = cellOptions.Style
				}
			}

			if styleToApply == nil && rowStyle != nil {
				styleToApply = rowStyle
			}

			if styleToApply == nil && column.Style != nil {
				styleToApply = column.Style
			}

			// Apply the determined style
			if err := t.applyCellStyle(styleToApply, actualColIndex, rowIndex, ops); err != nil {
				L().Warn("Failed to apply cell style",
					Int("column", actualColIndex),
					Int("row", rowIndex),
					Error(err))
				// Continue processing other cells even if one fails
				continue
			}
		}
	}

	return nil
}

// applyCellStyle applies a style configuration to a specific cell.
// If style is nil, no operation is performed.
func (t *Table) applyCellStyle(style *Style, colIndex, rowIndex int, ops TableOperations) error {
	if style == nil {
		return nil // No style to apply
	}

	// Apply the style using the operations interface
	if err := ops.applyStyleToCell(colIndex, rowIndex, *style); err != nil {
		return fmt.Errorf("failed to apply style to cell (%d,%d): %w", colIndex, rowIndex, err)
	}

	return nil
}

// applyColumnBorders applies borders to columns for all data rows.
// Handles both inner and boundary borders for each column.
func (t *Table) applyColumnBorders(dataStartRow, dataEndRow int, ops TableOperations) error {
	flatColumns := t.Columns.getFlattenedColumns()

	for colIndex, column := range flatColumns {
		actualColIndex := colIndex + 1

		// Skip columns without border configuration
		if column.Borders == nil || !column.Borders.hasBorders() {
			continue
		}

		// If inner borders are configured, apply to all cells in column
		if column.Borders.Inner != nil {
			for row := dataStartRow; row <= dataEndRow; row++ {
				if err := t.applyBordersToCell(actualColIndex, row, column.Borders.Inner, ops); err != nil {
					L().Warn("Failed to apply column border",
						Int("column", actualColIndex),
						Int("row", row),
						Error(err))
					continue
				}
			}
		} else {
			// Otherwise, apply left/right borders to all cells, top/bottom only to boundary cells
			for row := dataStartRow; row <= dataEndRow; row++ {
				cellBorder := &Borders{
					Left:  column.Borders.Left,
					Right: column.Borders.Right,
				}

				if row == dataStartRow {
					cellBorder.Top = column.Borders.Top
				}
				if row == dataEndRow {
					cellBorder.Bottom = column.Borders.Bottom
				}

				if err := t.applyBordersToCell(actualColIndex, row, cellBorder, ops); err != nil {
					L().Warn("Failed to apply column border",
						Int("column", actualColIndex),
						Int("row", row),
						Error(err))
					continue
				}
			}
		}
	}
	return nil
}

// applyRowBorders applies borders to all columns in a specific row.
// Handles both inner and boundary borders for each row.
func (t *Table) applyRowBorders(dataRowIndex, actualRowNum, totalColumns int, ops TableOperations) error {
	// Check if this row has a specific border configuration
	if rowOptions, exists := t.RowOptionsMap[dataRowIndex]; exists && rowOptions.Border != nil {
		// Skip if border config is empty
		if rowOptions.Border == nil || !rowOptions.Border.hasBorders() {
			return nil
		}

		// If inner borders are configured, apply to all columns in row
		if rowOptions.Border.Inner != nil {
			for col := 1; col <= totalColumns; col++ {
				if err := t.applyBordersToCell(col, actualRowNum, rowOptions.Border.Inner, ops); err != nil {
					L().Warn("Failed to apply row border",
						Int("column", col),
						Int("row", actualRowNum),
						Error(err))
					continue
				}
			}
		} else {
			// Otherwise, apply top/bottom borders to all cells, left/right only to boundary cells
			for col := 1; col <= totalColumns; col++ {
				cellBorders := &Borders{
					Top:    rowOptions.Border.Top,
					Bottom: rowOptions.Border.Bottom,
				}

				// Apply left border only to first column
				if col == 1 {
					cellBorders.Left = rowOptions.Border.Left
				}
				if col == totalColumns {
					cellBorders.Right = rowOptions.Border.Right
				}

				if err := t.applyBordersToCell(col, actualRowNum, cellBorders, ops); err != nil {
					L().Warn("Failed to apply row border",
						Int("column", col),
						Int("row", actualRowNum),
						Error(err))
					continue
				}
			}
		}
	}
	return nil
}

// applyCellSpecificBorders applies borders to individual cells based on cell configurations.
// Only cells with specific border configurations are processed.
func (t *Table) applyCellSpecificBorders(dataStartRow int, ops TableOperations) error {
	for colIndex, rowOptionsMap := range t.CellOptionsMap {
		for rowIndex, cellOptions := range rowOptionsMap {
			if cellOptions.Border != nil {
				actualRowNum := rowIndex + dataStartRow
				if err := t.applyBordersToCell(colIndex, actualRowNum, cellOptions.Border, ops); err != nil {
					L().Warn("Failed to apply cell-specific border",
						Int("column", colIndex),
						Int("row", actualRowNum),
						Error(err))
					continue
				}
			}
		}
	}
	return nil
}

// applyBordersToCell applies all configured borders to a specific cell.
// Each border (left, right, top, bottom) is applied if present.
func (t *Table) applyBordersToCell(col, row int, borders *Borders, ops TableOperations) error {
	if borders.Left != nil {
		if err := ops.applyBorderToCell(col, row, "left", borders.Left); err != nil {
			return fmt.Errorf("failed to apply left border: %w", err)
		}
	}
	if borders.Right != nil {
		if err := ops.applyBorderToCell(col, row, "right", borders.Right); err != nil {
			return fmt.Errorf("failed to apply right border: %w", err)
		}
	}
	if borders.Top != nil {
		if err := ops.applyBorderToCell(col, row, "top", borders.Top); err != nil {
			return fmt.Errorf("failed to apply top border: %w", err)
		}
	}
	if borders.Bottom != nil {
		if err := ops.applyBorderToCell(col, row, "bottom", borders.Bottom); err != nil {
			return fmt.Errorf("failed to apply bottom border: %w", err)
		}
	}
	return nil
}
