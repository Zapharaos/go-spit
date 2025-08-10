package table

import (
	"fmt"

	"github.com/Zapharaos/go-spit/internal/logger"
)

// RowConfigs maps row indices to their specific configurations.
type RowConfigs map[int]RowConfig

// RowConfig represents configuration options for a specific row in the export.
// This allows fine-grained control over individual rows, overriding default
// column-based settings when needed.
type RowConfig struct {
	RowIndex int           // The 0-based index of the row this configuration applies to
	Border   *BorderConfig // Optional border configuration for the entire row
	Merge    *MergeConfig  // Optional merge configuration that overrides column settings
	Style    *StyleConfig  // Optional style configuration for the entire row
}

// CellConfigs provides cell-level configuration mapping.
// The outer map keys are column indices, inner map keys are row indices.
type CellConfigs map[int]CellConfigsByRow

// CellConfigsByRow maps row indices to cell configurations for a specific column.
type CellConfigsByRow map[int]CellConfig

// CellConfig represents configuration options for a specific cell.
// This provides the finest level of control, allowing individual cells
// to override both column and row settings.
type CellConfig struct {
	RowIndex  int           // The 0-based row index of this cell
	ColIndex  int           // The 0-based column index of this cell
	Border    *BorderConfig // Optional border configuration for this cell
	Style     *StyleConfig  // Optional style configuration for this cell
	Mergeable bool          // Whether this cell can participate in merge operations
}

// BorderStyle represents the visual style of cell borders.
// These constants correspond to common border styles available in spreadsheet applications.
type BorderStyle int

const (
	BorderStyleNone   BorderStyle = 0 // No border
	BorderStyleThin   BorderStyle = 1 // Thin solid line
	BorderStyleMedium BorderStyle = 2 // Medium thickness solid line
	BorderStyleDashed BorderStyle = 3 // Dashed line
	BorderStyleDotted BorderStyle = 4 // Dotted line
	BorderStyleThick  BorderStyle = 5 // Thick solid line
	BorderStyleDouble BorderStyle = 6 // Double line
)

// BorderSide represents the configuration for one side of a cell border.
type BorderSide struct {
	Style BorderStyle // The visual style to apply to this border side
}

// BorderConfig represents complete border configuration for a cell, row, or column.
type BorderConfig struct {
	Left   *BorderSide   // Left border configuration
	Right  *BorderSide   // Right border configuration
	Top    *BorderSide   // Top border configuration
	Bottom *BorderSide   // Bottom border configuration
	Inner  *BorderConfig // Inner borders for ranges (used in some contexts)
}

// hasBorders checks if any borders are configured in this BorderConfig.
func (bc BorderConfig) hasBorders() bool {
	return (bc.Left != nil && bc.Left.Style != BorderStyleNone) ||
		(bc.Right != nil && bc.Right.Style != BorderStyleNone) ||
		(bc.Top != nil && bc.Top.Style != BorderStyleNone) ||
		(bc.Bottom != nil && bc.Bottom.Style != BorderStyleNone)
}

// SetInner creates inner border configuration with the same style for all sides.
func (bc BorderConfig) SetInner(style BorderStyle) BorderConfig {
	side := &BorderSide{Style: style}
	bc.Inner = &BorderConfig{
		Left:   side,
		Right:  side,
		Top:    side,
		Bottom: side,
	}
	return bc
}

// NewBorderConfig creates a BorderConfig with the same style applied to all sides.
func NewBorderConfig(style BorderStyle) BorderConfig {
	side := &BorderSide{Style: style}
	return BorderConfig{
		Left:   side,
		Right:  side,
		Top:    side,
		Bottom: side,
	}
}

// StyleConfig represents comprehensive styling configuration for cells.
type StyleConfig struct {
	Bold            bool          // Whether text should be bold
	Italic          bool          // Whether text should be italic
	Underline       string        // Underline style (format-specific values)
	TextColor       string        // Text color (usually hex format: "#RRGGBB")
	BackgroundColor string        // Cell background color (usually hex format: "#RRGGBB")
	FontSize        float64       // Font size in points
	FontFamily      string        // Font family name (e.g., "Arial", "Times New Roman")
	Alignment       CellAlignment // Text alignment within the cell
}

// CellAlignment represents the alignment options for cell content.
type CellAlignment int

const (
	AlignmentNone         CellAlignment = iota // Use default alignment
	AlignmentLeft                              // Left-aligned text, top-aligned vertically
	AlignmentCenter                            // Center-aligned text, top-aligned vertically
	AlignmentRight                             // Right-aligned text, top-aligned vertically
	AlignmentTop                               // Left-aligned text, top-aligned vertically
	AlignmentMiddle                            // Left-aligned text, center-aligned vertically
	AlignmentBottom                            // Left-aligned text, bottom-aligned vertically
	AlignmentCenterMiddle                      // Center-aligned text, center-aligned vertically
	AlignmentLeftMiddle                        // Left-aligned text, center-aligned vertically
	AlignmentRightMiddle                       // Right-aligned text, center-aligned vertically
)

// getAlignmentValues converts CellAlignment enum to horizontal and vertical alignment strings.
func (ca CellAlignment) getAlignmentValues() (horizontal, vertical string) {
	switch ca {
	case AlignmentLeft:
		return "left", "top"
	case AlignmentCenter:
		return "center", "top"
	case AlignmentRight:
		return "right", "top"
	case AlignmentTop:
		return "left", "top"
	case AlignmentMiddle:
		return "left", "center"
	case AlignmentBottom:
		return "left", "bottom"
	case AlignmentCenterMiddle:
		return "center", "center"
	case AlignmentLeftMiddle:
		return "left", "center"
	case AlignmentRightMiddle:
		return "right", "center"
	default:
		return "left", "top" // Default alignment for unspecified cases
	}
}

// RenderStyles applies styling and border operations to the table
func (t *Table) RenderStyles(ops Operations) error {
	dataStartRow := t.getDataStartRow()
	totalColumns := t.Columns.GetTotalColumnCount()
	dataEndRow := dataStartRow + len(t.Data) - 1

	// Phase 1: Apply header styles and borders
	if t.WriteHeader && len(t.Columns) > 0 {
		if err := t.applyHeaderStyles(ops); err != nil {
			return fmt.Errorf("failed to apply header styles: %w", err)
		}
	}

	// Phase 2: Apply data cell styles
	if err := t.applyCellStyles(dataStartRow, dataEndRow, ops); err != nil {
		return fmt.Errorf("failed to apply cell styles: %w", err)
	}

	// Phase 3: Apply column borders
	if err := t.applyColumnBorders(dataStartRow, dataEndRow, ops); err != nil {
		return fmt.Errorf("failed to apply column borders: %w", err)
	}

	// Phase 4: Apply row borders
	for rowIndex := range t.Data {
		actualRowNum := rowIndex + dataStartRow
		t.applyRowBorders(rowIndex, actualRowNum, totalColumns, ops)
	}

	// Phase 5: Apply cell-specific borders (applied last to override other border settings)
	if err := t.applyCellSpecificBorders(dataStartRow, ops); err != nil {
		return fmt.Errorf("failed to apply cell-specific borders: %w", err)
	}

	return nil
}

// applyHeaderStyles applies styling and borders to header rows
func (t *Table) applyHeaderStyles(ops Operations) error {
	maxDepth := t.Columns.GetMaxDepth()
	totalColumns := t.Columns.GetTotalColumnCount()

	// Apply bottom border to last header row
	lastHeaderRow := maxDepth
	for col := 1; col <= totalColumns; col++ {
		bottomBorder := &BorderSide{Style: BorderStyleThin}
		if err := ops.ApplyCellBorder(col, lastHeaderRow, "bottom", bottomBorder); err != nil {
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
func (t *Table) applyHeaderCellStyles(ops Operations) error {
	maxDepth := t.Columns.GetMaxDepth()
	totalColumns := t.Columns.GetTotalColumnCount()

	// Apply default header styling to all header cells
	headerStyle := StyleConfig{
		Bold:            true,
		BackgroundColor: "#E0E0E0",
		Alignment:       AlignmentCenter,
	}

	// Apply header styling to all header rows
	if err := ops.ApplyRangeStyle(1, 1, totalColumns, maxDepth, headerStyle); err != nil {
		logger.L().Warn("Failed to apply header range style", logger.Error(err))

		// Fallback: apply styles to individual header cells
		for row := 1; row <= maxDepth; row++ {
			for col := 1; col <= totalColumns; col++ {
				if err := t.applyCellStyle(&headerStyle, col, row, ops); err != nil {
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
func (t *Table) applyCellStyles(dataStartRow, dataEndRow int, ops Operations) error {
	flatColumns := t.Columns.GetFlattenedColumns()

	// Apply styles to each data row
	for rowIndex := dataStartRow; rowIndex <= dataEndRow; rowIndex++ {
		dataRowIndex := t.getDataIndexFromRowIndex(rowIndex)

		// Check if this data row exists (handle case where we have fewer data rows than expected)
		if dataRowIndex >= len(t.Data) {
			break
		}

		// Get row-level style if configured
		var rowStyle *StyleConfig
		if rc, exists := t.RowConfigs[dataRowIndex]; exists && rc.Style != nil {
			rowStyle = rc.Style
		}

		// Process each column in this row
		for colIndex, column := range flatColumns {
			actualColIndex := colIndex + 1
			var styleToApply *StyleConfig

			// Style priority: Cell-specific > Row-specific > Column-specific
			if cc, exists := t.CellConfigs[actualColIndex]; exists {
				if cellConfig, cellExists := cc[dataRowIndex]; cellExists && cellConfig.Style != nil {
					styleToApply = cellConfig.Style
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
func (t *Table) applyCellStyle(styleConfig *StyleConfig, colIndex, rowIndex int, ops Operations) error {
	if styleConfig == nil {
		return nil // No style to apply
	}

	// Apply the style using the operations interface
	if err := ops.ApplyCellStyle(colIndex, rowIndex, *styleConfig); err != nil {
		return fmt.Errorf("failed to apply style to cell (%d,%d): %w", colIndex, rowIndex, err)
	}

	return nil
}

// applyColumnBorders applies borders to columns for all data rows
func (t *Table) applyColumnBorders(dataStartRow, dataEndRow int, ops Operations) error {
	flatColumns := t.Columns.GetFlattenedColumns()

	for colIndex, column := range flatColumns {
		actualColIndex := colIndex + 1

		// Skip columns without border configuration
		if t.isEmptyBorderConfig(column.Border) {
			continue
		}

		// Check if this column has inner border configuration
		if column.Border.Inner != nil {
			// Apply complete borders to all cells in this column (inner borders enabled)
			for row := dataStartRow; row <= dataEndRow; row++ {
				if err := t.applyBorderConfigToCell(actualColIndex, row, column.Border, ops); err != nil {
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
				cellBorder := BorderConfig{
					Left:  column.Border.Left,  // Always apply left border
					Right: column.Border.Right, // Always apply right border
				}

				// Apply top border only to first row
				if row == dataStartRow {
					cellBorder.Top = column.Border.Top
				}

				// Apply bottom border only to last row
				if row == dataEndRow {
					cellBorder.Bottom = column.Border.Bottom
				}

				if err := t.applyBorderConfigToCell(actualColIndex, row, cellBorder, ops); err != nil {
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
func (t *Table) applyRowBorders(dataRowIndex, actualRowNum, totalColumns int, ops Operations) error {
	// Check if this row has a specific border configuration
	if rowConfig, exists := t.RowConfigs[dataRowIndex]; exists && rowConfig.Border != nil {
		// Skip if border config is empty
		if t.isEmptyBorderConfig(*rowConfig.Border) {
			return nil
		}

		// Check if this row has inner border configuration
		if rowConfig.Border.Inner != nil {
			// Apply complete borders to all columns in this row (inner borders enabled)
			for col := 1; col <= totalColumns; col++ {
				if err := t.applyBorderConfigToCell(col, actualRowNum, *rowConfig.Border, ops); err != nil {
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
				cellBorder := BorderConfig{
					Top:    rowConfig.Border.Top,    // Always apply top border
					Bottom: rowConfig.Border.Bottom, // Always apply bottom border
				}

				// Apply left border only to first column
				if col == 1 {
					cellBorder.Left = rowConfig.Border.Left
				}

				// Apply right border only to last column
				if col == totalColumns {
					cellBorder.Right = rowConfig.Border.Right
				}

				if err := t.applyBorderConfigToCell(col, actualRowNum, cellBorder, ops); err != nil {
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
func (t *Table) applyCellSpecificBorders(dataStartRow int, ops Operations) error {
	// Iterate through all cell-specific configurations
	for colIndex, rowConfigs := range t.CellConfigs {
		for rowIndex, cellConfig := range rowConfigs {
			// Only apply if this cell has a specific border configuration
			if cellConfig.Border != nil {
				actualRowNum := rowIndex + dataStartRow
				if err := t.applyBorderConfigToCell(colIndex, actualRowNum, *cellConfig.Border, ops); err != nil {
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

// applyBorderConfigToCell applies a complete border configuration to a specific cell
func (t *Table) applyBorderConfigToCell(col, row int, borderConfig BorderConfig, ops Operations) error {
	// Apply each border side if it's configured
	if borderConfig.Left != nil {
		if err := ops.ApplyCellBorder(col, row, "left", borderConfig.Left); err != nil {
			return fmt.Errorf("failed to apply left border: %w", err)
		}
	}

	if borderConfig.Right != nil {
		if err := ops.ApplyCellBorder(col, row, "right", borderConfig.Right); err != nil {
			return fmt.Errorf("failed to apply right border: %w", err)
		}
	}

	if borderConfig.Top != nil {
		if err := ops.ApplyCellBorder(col, row, "top", borderConfig.Top); err != nil {
			return fmt.Errorf("failed to apply top border: %w", err)
		}
	}

	if borderConfig.Bottom != nil {
		if err := ops.ApplyCellBorder(col, row, "bottom", borderConfig.Bottom); err != nil {
			return fmt.Errorf("failed to apply bottom border: %w", err)
		}
	}

	return nil
}

// isEmptyBorderConfig checks if a border configuration is effectively empty
func (t *Table) isEmptyBorderConfig(config BorderConfig) bool {
	return config.Left == nil && config.Right == nil && config.Top == nil && config.Bottom == nil && config.Inner == nil
}
