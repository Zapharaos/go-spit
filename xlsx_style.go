package go_spit

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

// autoFitColumns attempts to auto-fit column widths
func (xlsx XLSX) autoFitColumns() error {
	for i := 1; i <= xlsx.Columns.getTotalColumnCount(); i++ {
		colName, err := excelize.ColumnNumberToName(i)
		if err != nil {
			continue
		}

		// Set a reasonable default width
		if err = xlsx.File.SetColWidth(xlsx.SheetName, colName, colName, 15); err != nil {
			continue
		}
	}
	return nil
}

// applyCellStyles applies styling to all data cells after merging operations are complete
func (xlsx XLSX) applyCellStyles() error {
	// Calculate the starting row for data (accounting for header)
	dataStartRow := xlsx.Table.getDataStartRow()

	// Get flattened columns to match the data cell layout
	flatColumns := xlsx.Columns.getFlattenedColumns()

	// Apply styles to each data row
	for rowIndex := dataStartRow; rowIndex < dataStartRow+len(xlsx.Data); rowIndex++ {
		colIndex := 1

		dataRowIndex := xlsx.Table.getDataIndexFromRowIndex(rowIndex)

		// Check if row has configured a style
		var rowStyle *StyleConfig
		if rc, exists := xlsx.RowConfigs[dataRowIndex]; exists && rc.Style != nil {
			rowStyle = rc.Style
		}

		for _, column := range flatColumns {

			var style *StyleConfig

			// If cell has a specific style, use it
			if cc, exists := xlsx.CellConfigs[colIndex-1][dataRowIndex]; exists && cc.Style != nil {
				style = cc.Style
			} else if rowStyle != nil {
				// If row has a specific style, use it
				style = rowStyle
			} else {
				// Use column style, but check for merged cells and use parent style if needed
				style = column.Style

				// Use the XLSX struct directly as it now implements CellOperations
				if xlsx.IsCellMergedHorizontally(colIndex, rowIndex) {
					// If the cell is merged horizontally, try to find an unmerged parent column style
					// This is necessary to avoid applying styles to merged cells that span multiple columns
					// In this case it's alright if the merging is vertical since it's the same column
					// TODO : Implement
					/*if parentColumn := xslx.findUnmergedParentColumn(colIndex, rowIndex); parentColumn != nil {
						style = parentColumn.Style
					}*/
				}
			}

			if err := xlsx.applyCellStyle(style, colIndex, rowIndex); err != nil {
				return fmt.Errorf("failed to apply cell style at row %d, col %d: %w", rowIndex, colIndex, err)
			}
			colIndex++
		}
	}

	return nil
}

func (xlsx XLSX) applyCellStyle(styleConfig *StyleConfig, colIndex, rowIndex int) error {
	if styleConfig == nil {
		return nil // No style to apply
	}

	cellRef, err := excelize.CoordinatesToCellName(colIndex, rowIndex)
	if err != nil {
		return fmt.Errorf("failed to get cell reference: %w", err)
	}

	// Create the cell style
	var style *excelize.Style

	// Get existing style from the cell to preserve text formatting
	existingStyleID, err := xlsx.File.GetCellStyle(xlsx.SheetName, cellRef)
	if err != nil {
		style = &excelize.Style{}
	}

	// Get the existing style definition
	style, err = xlsx.File.GetStyle(existingStyleID)
	if err != nil {
		style = &excelize.Style{}
	}

	// Update the style with the provided configuration
	style.Font.Bold = styleConfig.Bold || style.Font.Bold
	style.Font.Italic = styleConfig.Italic || style.Font.Italic
	if styleConfig.Underline != "" {
		style.Font.Underline = styleConfig.Underline
	}
	if styleConfig.TextColor != "" {
		style.Font.Color = styleConfig.TextColor
	}
	if styleConfig.FontSize != 0 {
		style.Font.Size = styleConfig.FontSize
	}
	if styleConfig.FontFamily != "" {
		style.Font.Family = styleConfig.FontFamily
	}
	if styleConfig.BackgroundColor != "" {
		style.Fill = excelize.Fill{
			Type:    "pattern",
			Color:   []string{styleConfig.BackgroundColor},
			Pattern: 1,
		}
	}

	// Add alignment if specified
	if styleConfig.Alignment != AlignmentNone {
		horizontal, vertical := styleConfig.Alignment.getAlignmentValues()
		style.Alignment = &excelize.Alignment{
			Horizontal: horizontal,
			Vertical:   vertical,
		}
	}

	styleID, err := xlsx.File.NewStyle(style)
	if err != nil {
		return fmt.Errorf("failed to create cell style: %w", err)
	}

	if err = xlsx.File.SetCellStyle(xlsx.SheetName, cellRef, cellRef, styleID); err != nil {
		return fmt.Errorf("failed to set cell style: %w", err)
	}

	return nil
}
