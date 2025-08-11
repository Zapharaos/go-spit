// spreadsheet_excelize.go - Excelize-based implementation of the Spreadsheet interface.
//
// This file provides an adapter for the github.com/xuri/excelize library, enabling spreadsheet operations
// such as file creation, sheet management, cell formatting, and table manipulation in Excel files.

package go_spit

import (
	"fmt"
	"io"

	"github.com/xuri/excelize/v2"
)

// SpreadsheetExcelize provides Excelize-specific operations for spreadsheet handling.
// Implements the Spreadsheet interface using github.com/xuri/excelize.
type SpreadsheetExcelize struct {
	File      *excelize.File // Underlying Excelize file object
	SheetName string         // Current sheet name
	Table     *TableExcelize // Table adapter for Excelize
}

// NewSpreadsheetExcelize creates a new SpreadsheetExcelize instance for a given file, sheet, and table.
func NewSpreadsheetExcelize(file *excelize.File, sheetName string, t *Table) *SpreadsheetExcelize {
	return &SpreadsheetExcelize{
		File:      file,
		SheetName: sheetName,
		Table:     NewTableExcelize(file, sheetName, t),
	}
}

func (e *SpreadsheetExcelize) getTable() *Table {
	return e.Table.getTable()
}

func (e *SpreadsheetExcelize) getFile() interface{} {
	return e.File
}

func (e *SpreadsheetExcelize) createNewFile() error {
	e.File = excelize.NewFile()
	return nil
}

func (e *SpreadsheetExcelize) saveToWriter(writer io.Writer) error {
	_, err := e.File.WriteTo(writer)
	return err
}

func (e *SpreadsheetExcelize) close() error {
	return e.File.Close()
}

func (e *SpreadsheetExcelize) getSheetName() string {
	return e.SheetName
}

func (e *SpreadsheetExcelize) setSheetName(name string) {
	e.SheetName = name
}

func (e *SpreadsheetExcelize) createSheet() error {
	index, err := e.File.GetSheetIndex(e.SheetName)
	if err != nil || index == -1 {
		_, err = e.File.NewSheet(e.SheetName)
		if err != nil {
			return fmt.Errorf("failed to create sheet: %w", err)
		}
	}
	return nil
}

func (e *SpreadsheetExcelize) setActiveSheet() error {
	index, err := e.File.GetSheetIndex(e.SheetName)
	if err != nil {
		return fmt.Errorf("failed to get sheet index: %w", err)
	}
	e.File.SetActiveSheet(index)
	return nil
}

func (e *SpreadsheetExcelize) setColumnWidth(colLetter string, width float64) error {
	return e.File.SetColWidth(e.SheetName, colLetter, colLetter, width)
}

// Delegation to Excelize table operations
// These methods delegate to the TableExcelize adapter for cell and range operations.

func (e *SpreadsheetExcelize) getCellValue(col, row int) (string, error) {
	return e.Table.getCellValue(col, row)
}

func (e *SpreadsheetExcelize) setCellValue(col, row int, value interface{}) error {
	return e.Table.setCellValue(col, row, value)
}

func (e *SpreadsheetExcelize) mergeCells(startCol, startRow, endCol, endRow int) error {
	return e.Table.mergeCells(startCol, startRow, endCol, endRow)
}

func (e *SpreadsheetExcelize) isCellMerged(col, row int) bool {
	return e.Table.isCellMerged(col, row)
}

func (e *SpreadsheetExcelize) isCellMergedHorizontally(col, row int) bool {
	return e.Table.isCellMergedHorizontally(col, row)
}

func (e *SpreadsheetExcelize) applyBorderToCell(col, row int, side string, border *Border) error {
	return e.Table.applyBorderToCell(col, row, side, border)
}

func (e *SpreadsheetExcelize) applyBordersToRange(startCol, startRow, endCol, endRow int, borders Borders) error {
	return e.Table.applyBordersToRange(startCol, startRow, endCol, endRow, borders)
}

func (e *SpreadsheetExcelize) hasExistingBorder(col, row int, side string) bool {
	return e.Table.hasExistingBorder(col, row, side)
}

func (e *SpreadsheetExcelize) applyStyleToCell(col, row int, style Style) error {
	return e.Table.applyStyleToCell(col, row, style)
}

func (e *SpreadsheetExcelize) applyStyleToRange(startCol, startRow, endCol, endRow int, style Style) error {
	return e.Table.applyStyleToRange(startCol, startRow, endCol, endRow, style)
}

func (e *SpreadsheetExcelize) getColumnLetter(col int) string {
	return e.Table.getColumnLetter(col)
}

func (e *SpreadsheetExcelize) processValue(value interface{}, format string) (interface{}, error) {
	return e.Table.processValue(value, format)
}
