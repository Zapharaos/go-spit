// table_logic_style.go - Table cell styling logic.
//
// This file implements algorithms and operations for applying styles, borders, and formatting to exported tables.
// Styling errors are logged as warnings, allowing export to continue with best-effort visual enhancements.
// Used by spreadsheet backends to support advanced styling features in exported files.

package go_spit

import (
	"fmt"

	"github.com/Zapharaos/go-spit/internal/logger"
)

// renderStyles applies styling and border operations to the table
func (t *Table) renderStyles(ops tableOperations) error {
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

	// Apply row borders
	for rowIndex := range t.Data {
		actualRowNum := rowIndex + dataStartRow
		err := t.applyRowBorders(rowIndex, actualRowNum, totalColumns, ops)
		if err != nil {
			return fmt.Errorf("failed to apply row borders: %w", err)
		}
	}

	// Apply cell-specific borders (applied last to override other border settings)
	if err := t.applyCellSpecificBorders(dataStartRow, ops); err != nil {
		return fmt.Errorf("failed to apply cell-specific borders: %w", err)
	}

	return nil
}

// applyHeaderStyles applies styling and borders to header rows
func (t *Table) applyHeaderStyles(ops tableOperations) error {
	maxDepth := t.Columns.getMaxDepth()
	totalColumns := t.Columns.getTotalColumnCount()

	// Apply bottom border to last header row
	lastHeaderRow := maxDepth
	for col := 1; col <= totalColumns; col++ {
		bottomBorder := &Border{Style: BorderStyleThin}
		if err := ops.applyBorderToCell(col, lastHeaderRow, "bottom", bottomBorder); err != nil {
			logger.L().Warn("Failed to apply header bottom border",
				logger.Int("column", col),
				logger.Int("row", lastHeaderRow),
				logger.Error(err))
		}
	}

	// Apply header cell styles
	if err := t.applyHeaderCellStyles(ops); err != nil {
		return fmt.Errorf("failed to apply header cell styles: %w", err)
	}

	return nil
}

// applyHeaderCellStyles applies styling to header cells
func (t *Table) applyHeaderCellStyles(ops tableOperations) error {
	maxDepth := t.Columns.getMaxDepth()
	totalColumns := t.Columns.getTotalColumnCount()

	// Apply default header styling to all header cells
	headerStyle := Style{
		Bold:            true,
		BackgroundColor: "#E0E0E0",
		Alignment:       AlignmentCenter,
	}

	// Apply header styling to all header rows
	if err := ops.applyStyleToRange(1, 1, totalColumns, maxDepth, headerStyle); err != nil {
		logger.L().Warn("Failed to apply header range style", logger.Error(err))

		// Fallback: apply styles to individual header cells
		for row := 1; row <= maxDepth; row++ {
			for col := 1; col <= totalColumns; col++ {
				if err = t.applyCellStyle(&headerStyle, col, row, ops); err != nil {
					logger.L().Warn("Failed to apply header cell style",
						logger.Int("column", col),
						logger.Int("row", row),
						logger.Error(err))
				}
			}
		}
	}

	return nil
}

// applyCellStyles applies styling to all data cells based on priority: cell > row > column
func (t *Table) applyCellStyles(dataStartRow, dataEndRow int, ops tableOperations) error {
	flatColumns := t.Columns.getFlattenedColumns()

	// Apply styles to each data row
	for rowIndex := dataStartRow; rowIndex <= dataEndRow; rowIndex++ {
		dataRowIndex := t.getDataIndexFromRowIndex(rowIndex)

		// Check if this data row exists (handle case where we have fewer data rows than expected)
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
				logger.L().Warn("Failed to apply cell style",
					logger.Int("column", actualColIndex),
					logger.Int("row", rowIndex),
					logger.Error(err))
				// Continue processing other cells even if one fails
				continue
			}
		}
	}

	return nil
}

// applyCellStyle applies a style configuration to a specific cell
func (t *Table) applyCellStyle(style *Style, colIndex, rowIndex int, ops tableOperations) error {
	if style == nil {
		return nil // No style to apply
	}

	// Apply the style using the operations interface
	if err := ops.applyStyleToCell(colIndex, rowIndex, *style); err != nil {
		return fmt.Errorf("failed to apply style to cell (%d,%d): %w", colIndex, rowIndex, err)
	}

	return nil
}

// applyColumnBorders applies borders to columns for all data rows
func (t *Table) applyColumnBorders(dataStartRow, dataEndRow int, ops tableOperations) error {
	flatColumns := t.Columns.getFlattenedColumns()

	for colIndex, column := range flatColumns {
		actualColIndex := colIndex + 1

		// Skip columns without border configuration
		if !column.Borders.hasBorders() {
			continue
		}

		// Check if this column has inner border configuration
		if column.Borders.Inner != nil {
			// Apply complete borders to all cells in this column (inner borders enabled)
			for row := dataStartRow; row <= dataEndRow; row++ {
				if err := t.applyBordersToCell(actualColIndex, row, column.Borders, ops); err != nil {
					logger.L().Warn("Failed to apply column border",
						logger.Int("column", actualColIndex),
						logger.Int("row", row),
						logger.Error(err))
					continue
				}
			}
		} else {
			// Apply left/right borders to ALL cells, top/bottom only to boundary cells
			for row := dataStartRow; row <= dataEndRow; row++ {
				// Create border config for this specific cell
				cellBorder := Borders{
					Left:  column.Borders.Left,  // Always apply left border
					Right: column.Borders.Right, // Always apply right border
				}

				// Apply top border only to first row
				if row == dataStartRow {
					cellBorder.Top = column.Borders.Top
				}

				// Apply bottom border only to last row
				if row == dataEndRow {
					cellBorder.Bottom = column.Borders.Bottom
				}

				if err := t.applyBordersToCell(actualColIndex, row, cellBorder, ops); err != nil {
					logger.L().Warn("Failed to apply column border",
						logger.Int("column", actualColIndex),
						logger.Int("row", row),
						logger.Error(err))
					continue
				}
			}
		}
	}
	return nil
}

// applyRowBorders applies borders to all columns in a specific row
func (t *Table) applyRowBorders(dataRowIndex, actualRowNum, totalColumns int, ops tableOperations) error {
	// Check if this row has a specific border configuration
	if rowOptions, exists := t.RowOptionsMap[dataRowIndex]; exists && rowOptions.Border != nil {
		// Skip if border config is empty
		if !rowOptions.Border.hasBorders() {
			return nil
		}

		// Check if this row has inner border configuration
		if rowOptions.Border.Inner != nil {
			// Apply complete borders to all columns in this row (inner borders enabled)
			for col := 1; col <= totalColumns; col++ {
				if err := t.applyBordersToCell(col, actualRowNum, *rowOptions.Border, ops); err != nil {
					logger.L().Warn("Failed to apply row border",
						logger.Int("column", col),
						logger.Int("row", actualRowNum),
						logger.Error(err))
					continue
				}
			}
		} else {
			// Apply top/bottom borders to ALL cells, left/right only to boundary cells
			for col := 1; col <= totalColumns; col++ {
				// Create border config for this specific cell
				cellBorders := Borders{
					Top:    rowOptions.Border.Top,    // Always apply top border
					Bottom: rowOptions.Border.Bottom, // Always apply bottom border
				}

				// Apply left border only to first column
				if col == 1 {
					cellBorders.Left = rowOptions.Border.Left
				}

				// Apply right border only to last column
				if col == totalColumns {
					cellBorders.Right = rowOptions.Border.Right
				}

				if err := t.applyBordersToCell(col, actualRowNum, cellBorders, ops); err != nil {
					logger.L().Warn("Failed to apply row border",
						logger.Int("column", col),
						logger.Int("row", actualRowNum),
						logger.Error(err))
					continue
				}
			}
		}
	}
	return nil
}

// applyCellSpecificBorders applies borders to individual cells based on cell configurations
func (t *Table) applyCellSpecificBorders(dataStartRow int, ops tableOperations) error {
	// Iterate through all cell-specific configurations
	for colIndex, rowOptionsMap := range t.CellOptionsMap {
		for rowIndex, cellOptions := range rowOptionsMap {
			// Only apply if this cell has a specific border configuration
			if cellOptions.Border != nil {
				actualRowNum := rowIndex + dataStartRow
				if err := t.applyBordersToCell(colIndex, actualRowNum, *cellOptions.Border, ops); err != nil {
					logger.L().Warn("Failed to apply cell-specific border",
						logger.Int("column", colIndex),
						logger.Int("row", actualRowNum),
						logger.Error(err))
					// Continue processing other cells even if one fails
					continue
				}
			}
		}
	}
	return nil
}

// applyBordersToCell applies all configured borders to a specific cell
func (t *Table) applyBordersToCell(col, row int, borders Borders, ops tableOperations) error {
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
