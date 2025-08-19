// spreadsheet_excelize.go - Excelize-based implementation of the spreadsheet interface.
//
// This file provides an adapter for the github.com/xuri/excelize library, enabling spreadsheet operations
// such as file creation, sheet management, cell formatting, and table manipulation in Excel files.
// Implements the Spreadsheet interface for Excel file handling.

package spit

import (
	"fmt"
	"io"

	"github.com/xuri/excelize/v2"
)

// SpreadsheetExcelize provides Excelize-specific operations for spreadsheet handling.
// Implements the Spreadsheet interface using github.com/xuri/excelize.
type SpreadsheetExcelize struct {
	File      *excelize.File // Single Excelize file object for all sheets
	SheetName string         // Current sheet name
	Table     *TableExcelize // Current Table for Excelize
}

// NewSpreadsheetExcelize creates a new SpreadsheetExcelize instance for a given sheet name and table.
func NewSpreadsheetExcelize(sheetName string, t *Table) *SpreadsheetExcelize {
	return &SpreadsheetExcelize{
		SheetName: sheetName,
		Table:     NewTableExcelize(sheetName, t),
	}
}

// WithFile sets an existing Excelize file to the SpreadsheetExcelize instance.
// Keeps the TableExcelize adapter in sync with the spreadsheet file.
func (e *SpreadsheetExcelize) WithFile(file *excelize.File) *SpreadsheetExcelize {
	e.Table.WithFile(file) // Keep the TableExcelize in sync with the spreadsheet file
	e.File = file
	return e
}

// getTable returns the underlying Table object.
func (e *SpreadsheetExcelize) getTable() *Table {
	return e.Table.getTable()
}

// getFile returns the underlying Excelize file object.
func (e *SpreadsheetExcelize) getFile() interface{} {
	return e.File
}

// createNewFile initializes a new Excelize file and syncs it with the TableExcelize adapter.
func (e *SpreadsheetExcelize) createNewFile() error {
	f := excelize.NewFile()
	e.WithFile(f)
	e.Table.WithFile(f)
	return nil
}

// saveToWriter writes the Excelize file to an io.Writer (e.g., file, buffer).
func (e *SpreadsheetExcelize) saveToWriter(writer io.Writer) error {
	_, err := e.File.WriteTo(writer)
	return err
}

// close releases resources associated with the Excelize file.
func (e *SpreadsheetExcelize) close() error {
	return e.File.Close()
}

// getSheetName returns the current sheet name.
func (e *SpreadsheetExcelize) getSheetName() string {
	return e.SheetName
}

// setSheetName sets the active sheet name.
func (e *SpreadsheetExcelize) setSheetName(name string) {
	e.SheetName = name
}

// createSheet creates a new sheet with the current sheet name if it does not already exist.
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

// setActiveSheet sets the active sheet for subsequent operations.
func (e *SpreadsheetExcelize) setActiveSheet() error {
	index, err := e.File.GetSheetIndex(e.SheetName)
	if err != nil {
		return fmt.Errorf("failed to get sheet index: %w", err)
	}
	e.File.SetActiveSheet(index)
	return nil
}

// setColumnWidth sets the width of a column by its letter (e.g., "A", "B").
func (e *SpreadsheetExcelize) setColumnWidth(colLetter string, width float64) error {
	return e.File.SetColWidth(e.SheetName, colLetter, colLetter, width)
}

// Delegation to Excelize table operations
// These methods delegate to the TableExcelize adapter for cell and range operations.

// getCellValue returns the value of a cell at the given column and row.
func (e *SpreadsheetExcelize) getCellValue(col, row int) (string, error) {
	return e.Table.getCellValue(col, row)
}

// setCellValue sets the value of a cell at the given column and row.
func (e *SpreadsheetExcelize) setCellValue(col, row int, value interface{}) error {
	return e.Table.setCellValue(col, row, value)
}

// mergeCells merges a range of cells from startCol/startRow to endCol/endRow.
func (e *SpreadsheetExcelize) mergeCells(startCol, startRow, endCol, endRow int) error {
	return e.Table.mergeCells(startCol, startRow, endCol, endRow)
}

// isCellMerged checks if a cell is part of a merged range.
func (e *SpreadsheetExcelize) isCellMerged(col, row int) bool {
	return e.Table.isCellMerged(col, row)
}

// isCellMergedHorizontally checks if a cell is merged horizontally.
func (e *SpreadsheetExcelize) isCellMergedHorizontally(col, row int) bool {
	return e.Table.isCellMergedHorizontally(col, row)
}

// applyBorderToCell applies a border to a specific side of a cell.
func (e *SpreadsheetExcelize) applyBorderToCell(col, row int, side string, border *Border) error {
	return e.Table.applyBorderToCell(col, row, side, border)
}

// applyBordersToRange applies borders to a range of cells.
func (e *SpreadsheetExcelize) applyBordersToRange(startCol, startRow, endCol, endRow int, borders Borders) error {
	return e.Table.applyBordersToRange(startCol, startRow, endCol, endRow, borders)
}

// hasExistingBorder checks if a cell already has a border on a specific side.
func (e *SpreadsheetExcelize) hasExistingBorder(col, row int, side string) bool {
	return e.Table.hasExistingBorder(col, row, side)
}

// applyStyleToCell applies a style to a specific cell.
func (e *SpreadsheetExcelize) applyStyleToCell(col, row int, style Style) error {
	return e.Table.applyStyleToCell(col, row, style)
}

// applyStyleToRange applies a style to a range of cells.
func (e *SpreadsheetExcelize) applyStyleToRange(startCol, startRow, endCol, endRow int, style Style) error {
	return e.Table.applyStyleToRange(startCol, startRow, endCol, endRow, style)
}

// getColumnLetter returns the Excel column letter for a given column index.
func (e *SpreadsheetExcelize) getColumnLetter(col int) string {
	return e.Table.getColumnLetter(col)
}

// processValue processes a value according to the specified format for Excel output.
func (e *SpreadsheetExcelize) processValue(value interface{}, format string) (interface{}, error) {
	return e.Table.processValue(value, format)
}
