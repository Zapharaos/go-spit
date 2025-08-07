package go_spit

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

// Ensure XLSX implements the CellOperations interface
var _ CellOperations = (*XLSX)(nil)

// GetCellValue retrieves the value from a cell at the specified coordinates
func (xlsx XLSX) GetCellValue(col, row int) (string, error) {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return "", err
	}
	return xlsx.File.GetCellValue(xlsx.SheetName, cellRef)
}

// SetCellValue sets a value in a cell at the specified coordinates
func (xlsx XLSX) SetCellValue(col, row int, value interface{}) error {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return err
	}
	return xlsx.File.SetCellValue(xlsx.SheetName, cellRef, value)
}

// MergeCells merges cells from (startCol, startRow) to (endCol, endRow)
func (xlsx XLSX) MergeCells(startCol, startRow, endCol, endRow int) error {
	startCell, err1 := excelize.CoordinatesToCellName(startCol, startRow)
	endCell, err2 := excelize.CoordinatesToCellName(endCol, endRow)
	if err1 != nil || err2 != nil {
		return fmt.Errorf("failed to convert coordinates to cell names: %v, %v", err1, err2)
	}
	return xlsx.File.MergeCell(xlsx.SheetName, startCell, endCell)
}

// IsCellMerged checks if a cell is part of a merged range
func (xlsx XLSX) IsCellMerged(col, row int) bool {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return false
	}
	return xlsx.isCellMerged(cellRef)
}

// IsCellMergedHorizontally checks if a cell is part of a horizontally merged range
func (xlsx XLSX) IsCellMergedHorizontally(col, row int) bool {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return false
	}
	return xlsx.isCellMergedHorizontally(cellRef)
}

// GetColumnLetter converts column index to Excel column letter
func (xlsx XLSX) GetColumnLetter(col int) string {
	colLetter, err := excelize.ColumnNumberToName(col)
	if err != nil {
		return "A" // Fallback to column A if conversion fails
	}
	return colLetter
}

// ProcessValue processes a raw value according to a format specification
func (xlsx XLSX) ProcessValue(value interface{}, format string) (interface{}, error) {
	// Delegate to the existing processValueForCell method
	return xlsx.processValueForCell(value, format)
}

// isCellMerged checks if a cell is part of a merged range
func (xlsx XLSX) isCellMerged(cellRef string) bool {
	mergedCells, err := xlsx.File.GetMergeCells(xlsx.SheetName)
	if err != nil {
		return false
	}

	col, row, err := excelize.CellNameToCoordinates(cellRef)
	if err != nil {
		return false
	}

	for _, mergedCell := range mergedCells {
		startCol, startRow, err1 := excelize.CellNameToCoordinates(mergedCell.GetStartAxis())
		endCol, endRow, err2 := excelize.CellNameToCoordinates(mergedCell.GetEndAxis())

		if err1 == nil && err2 == nil {
			if col >= startCol && col <= endCol && row >= startRow && row <= endRow {
				return true
			}
		}
	}
	return false
}

// isCellMergedHorizontally checks if a cell is part of a horizontally merged range (spans multiple columns)
func (xlsx XLSX) isCellMergedHorizontally(cellRef string) bool {
	mergedCells, err := xlsx.File.GetMergeCells(xlsx.SheetName)
	if err != nil {
		return false
	}

	col, row, err := excelize.CellNameToCoordinates(cellRef)
	if err != nil {
		return false
	}

	for _, mergedCell := range mergedCells {
		startCol, startRow, err1 := excelize.CellNameToCoordinates(mergedCell.GetStartAxis())
		endCol, endRow, err2 := excelize.CellNameToCoordinates(mergedCell.GetEndAxis())

		if err1 == nil && err2 == nil {
			if col >= startCol && col <= endCol && row >= startRow && row <= endRow {
				if endCol > startCol { // merged horizontally
					return true
				}
			}
		}
	}
	return false
}
