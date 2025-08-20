// excelize_table.go - Excelize-based implementation of TableOperations for go-spit
//
// This file provides an adapter for the github.com/xuri/excelize library, enabling table operations
// such as cell access, merging, styling, and value formatting for Excel spreadsheets.

package spit

import (
	"fmt"
	"time"

	"github.com/xuri/excelize/v2"
)

// TableExcelize provides Excelize-specific operations for table handling.
// Implements TableOperations for Excel spreadsheets using github.com/xuri/excelize.
type TableExcelize struct {
	File      *excelize.File // Underlying Excelize file object
	SheetName string         // Current sheet name
	Table     *Table         // Reference to the generic Table struct
}

// NewTableExcelize creates a new TableExcelize instance for a given sheet name and table.
// The Excelize file is optional, set later via WithFile.
func NewTableExcelize(sheetName string, table *Table) *TableExcelize {
	return &TableExcelize{
		SheetName: sheetName,
		Table:     table,
	}
}

// WithFile sets an existing Excelize file to the TableExcelize instance.
// Returns the TableExcelize for chaining.
func (e *TableExcelize) WithFile(file *excelize.File) *TableExcelize {
	e.File = file
	return e
}

// getTable returns the underlying Table struct for direct access/manipulation.
func (e *TableExcelize) getTable() *Table {
	return e.Table
}

// getCellValue returns the value of a cell at the given column and row (1-based indices).
// Converts coordinates to Excel cell reference and retrieves the value from the sheet.
func (e *TableExcelize) getCellValue(col, row int) (string, error) {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return "", err
	}
	return e.File.GetCellValue(e.SheetName, cellRef)
}

// setCellValue sets the value of a cell at the given column and row.
// Converts coordinates to Excel cell reference and sets the value in the sheet.
func (e *TableExcelize) setCellValue(col, row int, value interface{}) error {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return err
	}
	return e.File.SetCellValue(e.SheetName, cellRef, value)
}

// mergeCells merges a rectangular range of cells from start to end coordinates.
// Converts coordinates to Excel cell references and merges the specified range.
func (e *TableExcelize) mergeCells(startCol, startRow, endCol, endRow int) error {
	startCell, err1 := excelize.CoordinatesToCellName(startCol, startRow)
	endCell, err2 := excelize.CoordinatesToCellName(endCol, endRow)
	if err1 != nil || err2 != nil {
		return fmt.Errorf("failed to convert coordinates: %v, %v", err1, err2)
	}
	return e.File.MergeCell(e.SheetName, startCell, endCell)
}

// isCellMerged checks if a cell at the given column and row is merged with others.
// Returns true if the cell is part of a merged range, false otherwise.
func (e *TableExcelize) isCellMerged(col, row int) bool {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return false
	}

	mergedCells, err := e.File.GetMergeCells(e.SheetName)
	if err != nil {
		return false
	}
	for _, mergeCell := range mergedCells {
		if isCellInRange(cellRef, mergeCell.GetStartAxis(), mergeCell.GetEndAxis()) {
			return true
		}
	}
	return false
}

// isCellMergedHorizontally checks if a cell at the given column and row is merged horizontally.
// Returns true if the cell is part of a horizontally merged range, false otherwise.
func (e *TableExcelize) isCellMergedHorizontally(col, row int) bool {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return false
	}

	mergedCells, err := e.File.GetMergeCells(e.SheetName)
	if err != nil {
		return false
	}
	for _, mergeCell := range mergedCells {
		if isCellInRange(cellRef, mergeCell.GetStartAxis(), mergeCell.GetEndAxis()) {
			startCol, startRow, _ := excelize.CellNameToCoordinates(mergeCell.GetStartAxis())
			endCol, endRow, _ := excelize.CellNameToCoordinates(mergeCell.GetEndAxis())
			return startRow == endRow && startCol != endCol
		}
	}
	return false
}

// applyBorderToCell applies a border to a specific side of a cell at the given column and row.
// The border style is defined by the Border parameter. Only non-none styles are applied.
func (e *TableExcelize) applyBorderToCell(col, row int, side string, border *Border) error {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return err
	}
	if border == nil || border.Style == BorderStyleNone {
		return nil
	}

	// Get current style and preserve borders
	excelStyle, err := e.getCellStyle(col, row)
	if excelStyle == nil || err != nil {
		excelStyle = &excelize.Style{}
	} else {
		switch side {
		case "left", "right", "top", "bottom":
			excelStyle.Border = append(excelStyle.Border, excelize.Border{Type: side, Color: "000000", Style: int(border.Style)})
		default:
			return fmt.Errorf("unsupported border side: %s", side)
		}
	}

	styleID, err := e.File.NewStyle(excelStyle)
	if err != nil {
		return err
	}

	return e.File.SetCellStyle(e.SheetName, cellRef, cellRef, styleID)
}

// applyBordersToRange applies borders to a range of cells defined by start and end coordinates.
// Each side of the range can have a different border style, as specified in the Borders parameter.
func (e *TableExcelize) applyBordersToRange(startCol, startRow, endCol, endRow int, borders Borders) error {
	for row := startRow; row <= endRow; row++ {
		for col := startCol; col <= endCol; col++ {
			if col == startCol && borders.Left != nil {
				if err := e.applyBorderToCell(col, row, "left", borders.Left); err != nil {
					return err
				}
			}
			if col == endCol && borders.Right != nil {
				if err := e.applyBorderToCell(col, row, "right", borders.Right); err != nil {
					return err
				}
			}
			if row == startRow && borders.Top != nil {
				if err := e.applyBorderToCell(col, row, "top", borders.Top); err != nil {
					return err
				}
			}
			if row == endRow && borders.Bottom != nil {
				if err := e.applyBorderToCell(col, row, "bottom", borders.Bottom); err != nil {
					return err
				}
			}
		}
	}
	return nil
}

// hasExistingBorder checks if a cell at the given column and row has any existing border applied on the specified side.
// Returns true if there is a border style applied, false otherwise.
func (e *TableExcelize) hasExistingBorder(col, row int, side string) bool {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return false
	}
	styleID, err := e.File.GetCellStyle(e.SheetName, cellRef)
	if err != nil {
		return false
	}
	// Check if there's a style applied (simple check)
	return styleID > 0
}

// applyStyleToCell applies a style to a cell at the given column and row.
// The style properties are defined in the style parameter. Existing borders are preserved.
func (e *TableExcelize) applyStyleToCell(col, row int, style Style) error {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return err
	}

	inputStyle := convertStyleToExcelizeStyle(style)

	// Get current style and preserve borders
	excelStyle, err := e.getCellStyle(col, row)
	if excelStyle == nil || err != nil {
		// If no existing style, use the input style directly
		excelStyle = inputStyle
	} else {
		// Merge input style with existing style
		// Preserve existing style properties if the inputStyle has nil values in them

		if inputStyle.Fill.Color != nil {
			excelStyle.Fill = inputStyle.Fill
		}
		if inputStyle.Font != nil {
			excelStyle.Font = inputStyle.Font
		}
		if inputStyle.Alignment != nil {
			excelStyle.Alignment = inputStyle.Alignment
		}
	}

	styleID, err := e.File.NewStyle(excelStyle)
	if err != nil {
		return err
	}
	return e.File.SetCellStyle(e.SheetName, cellRef, cellRef, styleID)
}

// applyStyleToRange applies a style to a range of cells defined by start and end coordinates.
// The style properties are defined in the style parameter.
func (e *TableExcelize) applyStyleToRange(startCol, startRow, endCol, endRow int, style Style) error {
	for row := startRow; row <= endRow; row++ {
		for col := startCol; col <= endCol; col++ {
			if err := e.applyStyleToCell(col, row, style); err != nil {
				return err
			}
		}
	}
	return nil
}

func (e *TableExcelize) getCellStyle(col, row int) (*excelize.Style, error) {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return nil, err
	}
	styleID, err := e.File.GetCellStyle(e.SheetName, cellRef)
	if err != nil {
		return nil, err
	}
	return e.File.GetStyle(styleID)
}

// getColumnLetter returns the Excel-style column letter (A, B, C, ...) for a given column number.
func (e *TableExcelize) getColumnLetter(col int) string {
	letter, _ := excelize.ColumnNumberToName(col)
	return letter
}

// processValue processes and formats a value according to its type and the specified format.
// Supports basic types, time.Time, and slices. Formats value for Excel export.
func (e *TableExcelize) processValue(value interface{}, format string) (interface{}, error) {
	switch v := value.(type) {
	case []interface{}:
		if e.Table.ListSeparator != "" {
			return convertSliceToString(v, format, e.Table.ListSeparator)
		}
		return fmt.Sprintf("%v", v), nil
	case time.Time:
		if format != "" {
			return v.Format(format), nil
		}
		return v, nil
	case *time.Time:
		if v != nil {
			if format != "" {
				return v.Format(format), nil
			}
			return *v, nil
		}
		return "", nil
	case string:
		// Skip formatting for string values, even if format is specified
		// This prevents format conflicts (e.g., "Total" being formatted as date)
		return v, nil
	case int, int8, int16, int32, int64:
		return v, nil
	case uint, uint8, uint16, uint32, uint64:
		return v, nil
	case float32, float64:
		return v, nil
	case bool:
		return v, nil
	default:
		if format != "" {
			var err error
			value, err = formatValue(value, format)
			if err != nil {
				return "", err
			}
		}
		return fmt.Sprintf("%v", value), nil
	}
}

// convertStyleToExcelizeStyle converts a Style struct to the corresponding Excelize style.
// Maps font, fill, and alignment properties to Excelize style attributes.
func convertStyleToExcelizeStyle(style Style) *excelize.Style {
	excelStyle := &excelize.Style{}

	if style.Bold || style.Italic || style.FontSize > 0 || style.FontFamily != "" || style.TextColor != "" {
		font := &excelize.Font{}
		if style.Bold {
			font.Bold = true
		}
		if style.Italic {
			font.Italic = true
		}
		if style.FontSize > 0 {
			font.Size = style.FontSize
		}
		if style.FontFamily != "" {
			font.Family = style.FontFamily
		}
		if style.TextColor != "" {
			font.Color = style.TextColor
		}
		if style.Underline != "" {
			font.Underline = style.Underline
		}
		excelStyle.Font = font
	}

	if style.BackgroundColor != "" {
		excelStyle.Fill = excelize.Fill{
			Type:    "pattern",
			Color:   []string{style.BackgroundColor},
			Pattern: 1,
		}
	}

	if style.Alignment != AlignmentNone {
		horizontal, vertical := style.Alignment.getAlignmentValues()
		excelStyle.Alignment = &excelize.Alignment{
			Horizontal: horizontal,
			Vertical:   vertical,
		}
	}

	return excelStyle
}

// isCellInRange checks if a cell reference is within a given range defined by start and end references.
// Returns true if the cell is in range, false otherwise.
func isCellInRange(cellRef, startRef, endRef string) bool {
	col, row, err := excelize.CellNameToCoordinates(cellRef)
	if err != nil {
		return false
	}
	startCol, startRow, err := excelize.CellNameToCoordinates(startRef)
	if err != nil {
		return false
	}
	endCol, endRow, err := excelize.CellNameToCoordinates(endRef)
	if err != nil {
		return false
	}
	return col >= startCol && col <= endCol && row >= startRow && row <= endRow
}
