// excelize_spreadsheet.go - Excelize-based implementation of the spreadsheet interface.
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

// GetTable returns the underlying Table object.
func (e *SpreadsheetExcelize) GetTable() *Table {
	return e.Table.GetTable()
}

// GetFile returns the underlying Excelize file object.
func (e *SpreadsheetExcelize) GetFile() interface{} {
	return e.File
}

// CreateNewFile initializes a new Excelize file and syncs it with the TableExcelize adapter.
func (e *SpreadsheetExcelize) CreateNewFile() error {
	f := excelize.NewFile()
	e.WithFile(f)
	e.Table.WithFile(f)
	return nil
}

// SaveToWriter writes the Excelize file to an io.Writer (e.g., file, buffer).
func (e *SpreadsheetExcelize) SaveToWriter(writer io.Writer) error {
	_, err := e.File.WriteTo(writer)
	return err
}

// Close releases resources associated with the Excelize file.
func (e *SpreadsheetExcelize) Close() error {
	return e.File.Close()
}

// GetSheetName returns the current sheet name.
func (e *SpreadsheetExcelize) GetSheetName() string {
	return e.SheetName
}

// SetSheetName sets the active sheet name.
func (e *SpreadsheetExcelize) SetSheetName(name string) {
	e.SheetName = name
}

// CreateSheet creates a new sheet with the current sheet name if it does not already exist.
func (e *SpreadsheetExcelize) CreateSheet() error {
	index, err := e.File.GetSheetIndex(e.SheetName)
	if err != nil || index == -1 {
		_, err = e.File.NewSheet(e.SheetName)
		if err != nil {
			return fmt.Errorf("failed to create sheet: %w", err)
		}
	}
	return nil
}

// SetActiveSheet sets the active sheet for subsequent operations.
func (e *SpreadsheetExcelize) SetActiveSheet() error {
	index, err := e.File.GetSheetIndex(e.SheetName)
	if err != nil {
		return fmt.Errorf("failed to get sheet index: %w", err)
	}
	e.File.SetActiveSheet(index)
	return nil
}

// SetColumnWidth sets the width of a column by its letter (e.g., "A", "B").
func (e *SpreadsheetExcelize) SetColumnWidth(colLetter string, width float64) error {
	return e.File.SetColWidth(e.SheetName, colLetter, colLetter, width)
}

// Delegation to Excelize table operations
// These methods delegate to the TableExcelize adapter for cell and range operations.

// GetCellValue returns the value of a cell at the given column and row.
func (e *SpreadsheetExcelize) GetCellValue(col, row int) (string, error) {
	return e.Table.GetCellValue(col, row)
}

// SetCellValue sets the value of a cell at the given column and row.
func (e *SpreadsheetExcelize) SetCellValue(col, row int, value interface{}) error {
	return e.Table.SetCellValue(col, row, value)
}

// MergeCells merges a range of cells from startCol/startRow to endCol/endRow.
func (e *SpreadsheetExcelize) MergeCells(startCol, startRow, endCol, endRow int) error {
	return e.Table.MergeCells(startCol, startRow, endCol, endRow)
}

// IsCellMerged checks if a cell is part of a merged range.
func (e *SpreadsheetExcelize) IsCellMerged(col, row int) bool {
	return e.Table.IsCellMerged(col, row)
}

// IsCellMergedHorizontally checks if a cell is merged horizontally.
func (e *SpreadsheetExcelize) IsCellMergedHorizontally(col, row int) bool {
	return e.Table.IsCellMergedHorizontally(col, row)
}

// ApplyBorderToCell applies a border to a specific side of a cell.
func (e *SpreadsheetExcelize) ApplyBorderToCell(col, row int, side string, border *Border) error {
	return e.Table.ApplyBorderToCell(col, row, side, border)
}

// ApplyBordersToRange applies borders to a range of cells.
func (e *SpreadsheetExcelize) ApplyBordersToRange(startCol, startRow, endCol, endRow int, borders Borders) error {
	return e.Table.ApplyBordersToRange(startCol, startRow, endCol, endRow, borders)
}

// HasExistingBorder checks if a cell already has a border on a specific side.
func (e *SpreadsheetExcelize) HasExistingBorder(col, row int, side string) bool {
	return e.Table.HasExistingBorder(col, row, side)
}

// ApplyStyleToCell applies a style to a specific cell.
func (e *SpreadsheetExcelize) ApplyStyleToCell(col, row int, style Style) error {
	return e.Table.ApplyStyleToCell(col, row, style)
}

// ApplyStyleToRange applies a style to a range of cells.
func (e *SpreadsheetExcelize) ApplyStyleToRange(startCol, startRow, endCol, endRow int, style Style) error {
	return e.Table.ApplyStyleToRange(startCol, startRow, endCol, endRow, style)
}

// GetColumnLetter returns the Excel column letter for a given column index.
func (e *SpreadsheetExcelize) GetColumnLetter(col int) string {
	return e.Table.GetColumnLetter(col)
}

// ProcessValue processes a value according to the specified format for Excel output.
func (e *SpreadsheetExcelize) ProcessValue(value interface{}, format string) (interface{}, error) {
	return e.Table.ProcessValue(value, format)
}
