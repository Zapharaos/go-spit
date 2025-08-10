// table_excelize.go - Excelize-based implementation of tableOperations for go-spit
//
// This file provides an adapter for the github.com/xuri/excelize library, enabling table operations
// such as cell access, merging, styling, and value formatting for Excel spreadsheets.

package go_spit

import (
	"fmt"
	"time"

	"github.com/xuri/excelize/v2"
)

// TableExcelize provides Excelize-specific operations for table handling.
// Implements tableOperations for Excel spreadsheets using github.com/xuri/excelize.
type TableExcelize struct {
	File      *excelize.File // Underlying Excelize file object
	SheetName string         // Current sheet name
	Table     *Table         // Reference to the generic Table struct
}

// NewTableExcelize creates a new TableExcelize instance for a given file, sheet, and table.
func NewTableExcelize(file *excelize.File, sheetName string, table *Table) *TableExcelize {
	return &TableExcelize{
		File:      file,
		SheetName: sheetName,
		Table:     table,
	}
}

// GetTable returns the underlying Table struct for direct access/manipulation.
func (e *TableExcelize) getTable() *Table {
	return e.Table
}

// GetCellValue returns the value of a cell at the given column and row (1-based indices).
// Converts coordinates to Excel cell reference and retrieves the value from the sheet.
func (e *TableExcelize) getCellValue(col, row int) (string, error) {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return "", err
	}
	return e.File.GetCellValue(e.SheetName, cellRef)
}

// SetCellValue sets the value of a cell at the given column and row.
// Converts coordinates to Excel cell reference and sets the value in the sheet.
func (e *TableExcelize) setCellValue(col, row int, value interface{}) error {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return err
	}
	return e.File.SetCellValue(e.SheetName, cellRef, value)
}

// MergeCells merges a rectangular range of cells from start to end coordinates.
// Converts coordinates to Excel cell references and merges the specified range.
func (e *TableExcelize) mergeCells(startCol, startRow, endCol, endRow int) error {
	startCell, err1 := excelize.CoordinatesToCellName(startCol, startRow)
	endCell, err2 := excelize.CoordinatesToCellName(endCol, endRow)
	if err1 != nil || err2 != nil {
		return fmt.Errorf("failed to convert coordinates: %v, %v", err1, err2)
	}
	return e.File.MergeCell(e.SheetName, startCell, endCell)
}

// IsCellMerged checks if a cell at the given column and row is merged with others.
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

// IsCellMergedHorizontally checks if a cell at the given column and row is merged horizontally.
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

// ApplyBorderToCell applies a border to a specific side of a cell at the given column and row.
// The border style is defined by the Border parameter.
func (e *TableExcelize) applyBorderToCell(col, row int, side string, border *Border) error {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return err
	}
	return e.applyCellBorder(cellRef, side, border)
}

// ApplyBorderToRange applies borders to a range of cells defined by start and end coordinates.
// Each side of the range can have a different border style, as specified in the Borders parameter.
func (e *TableExcelize) applyBorderToRange(startCol, startRow, endCol, endRow int, borders Borders) error {
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

// HasExistingBorder checks if a cell at the given column and row has any existing border applied.
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

// ApplyStyleToCell applies a style to a cell at the given column and row.
// The style properties are defined in the style parameter.
func (e *TableExcelize) applyStyleToCell(col, row int, style Style) error {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return err
	}
	return e.applyCellStyle(cellRef, style)
}

// ApplyStyleToRange applies a style to a range of cells defined by start and end coordinates.
// The style properties are defined in the style parameter.
func (e *TableExcelize) applyStyleToRange(startCol, startRow, endCol, endRow int, style Style) error {
	startCell, err1 := excelize.CoordinatesToCellName(startCol, startRow)
	endCell, err2 := excelize.CoordinatesToCellName(endCol, endRow)
	if err1 != nil || err2 != nil {
		return fmt.Errorf("failed to convert coordinates: %v, %v", err1, err2)
	}

	excelStyle := convertStyleToExcelizeStyle(style)
	styleID, err := e.File.NewStyle(excelStyle)
	if err != nil {
		return err
	}

	return e.File.SetCellStyle(e.SheetName, startCell, endCell, styleID)
}

// GetColumnLetter returns the Excel-style column letter (A, B, C, ...) for a given column number.
func (e *TableExcelize) getColumnLetter(col int) string {
	letter, _ := excelize.ColumnNumberToName(col)
	return letter
}

// ProcessValue processes and formats a value according to its type and the specified format.
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

// applyCellBorder applies a border to a cell reference for a specific side (left, right, top, bottom).
// The border style is determined by the Border parameter.
func (e *TableExcelize) applyCellBorder(cellRef, borderType string, border *Border) error {
	if border == nil || border.Style == BorderStyleNone {
		return nil
	}

	excelBorderStyle := convertBorderStyleToExcelize(border.Style)

	// Create the style with border
	style := &excelize.Style{}

	switch borderType {
	case "left":
		style.Border = []excelize.Border{{Type: "left", Color: "000000", Style: excelBorderStyle}}
	case "right":
		style.Border = []excelize.Border{{Type: "right", Color: "000000", Style: excelBorderStyle}}
	case "top":
		style.Border = []excelize.Border{{Type: "top", Color: "000000", Style: excelBorderStyle}}
	case "bottom":
		style.Border = []excelize.Border{{Type: "bottom", Color: "000000", Style: excelBorderStyle}}
	default:
		return fmt.Errorf("unsupported border type: %s", borderType)
	}

	styleID, err := e.File.NewStyle(style)
	if err != nil {
		return err
	}

	return e.File.SetCellStyle(e.SheetName, cellRef, cellRef, styleID)
}

// applyCellStyle applies a style to a cell reference based on the properties defined in the style parameter.
func (e *TableExcelize) applyCellStyle(cellRef string, style Style) error {
	excelStyle := convertStyleToExcelizeStyle(style)
	styleID, err := e.File.NewStyle(excelStyle)
	if err != nil {
		return err
	}
	return e.File.SetCellStyle(e.SheetName, cellRef, cellRef, styleID)
}

// convertBorderStyleToExcelize converts a BorderStyle enum value to the corresponding Excelize border style integer.
func convertBorderStyleToExcelize(style BorderStyle) int {
	switch style {
	case BorderStyleNone:
		return 0
	case BorderStyleThin:
		return 1
	case BorderStyleMedium:
		return 2
	case BorderStyleDashed:
		return 3
	case BorderStyleDotted:
		return 4
	case BorderStyleThick:
		return 5
	case BorderStyleDouble:
		return 6
	default:
		return 1
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

	// TODO : check function
	if style.Alignment != AlignmentNone {
		alignment := &excelize.Alignment{}
		switch style.Alignment {
		case AlignmentLeft:
			alignment.Horizontal = "left"
		case AlignmentCenter:
			alignment.Horizontal = "center"
		case AlignmentRight:
			alignment.Horizontal = "right"
		}
		excelStyle.Alignment = alignment
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
