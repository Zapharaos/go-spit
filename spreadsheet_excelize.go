// spreadsheet_excelize.go - Excelize-based implementation of the spreadsheet interface.
//
// This file provides an adapter for the github.com/xuri/excelize library, enabling spreadsheet operations
// such as file creation, sheet management, cell formatting, and table manipulation in Excel files.

package spit

import (
	"fmt"
	"io"

	"github.com/xuri/excelize/v2"
)

// SpreadsheetExcelize provides Excelize-specific operations for spreadsheet handling.
// Implements the Spreadsheet interface using github.com/xuri/excelize.
type SpreadsheetExcelize struct {
	File *excelize.File // Single Excelize file object for all sheets
	// Tables    map[string]*TableExcelize
	SheetName string         // Current sheet name
	Table     *TableExcelize // Current Table for Excelize
}

// NewSpreadsheetExcelize creates a new SpreadsheetExcelize instance for a given file, sheet, and table.
func NewSpreadsheetExcelize(sheetName string, t *Table) *SpreadsheetExcelize {
	return &SpreadsheetExcelize{
		SheetName: sheetName,
		Table:     NewTableExcelize(sheetName, t),
	}
}

// WithFile allows setting an existing Excelize file to the SpreadsheetExcelize instance.
func (e *SpreadsheetExcelize) WithFile(file *excelize.File) *SpreadsheetExcelize {
	e.Table.WithFile(file) // Keep the TableExcelize in sync with the spreadsheet file
	e.File = file
	return e
}

func (e *SpreadsheetExcelize) GetTable() *Table {
	return e.Table.GetTable()
}

func (e *SpreadsheetExcelize) GetFile() interface{} {
	return e.File
}

func (e *SpreadsheetExcelize) CreateNewFile() error {
	f := excelize.NewFile()
	e.WithFile(f)
	e.Table.WithFile(f)
	return nil
}

func (e *SpreadsheetExcelize) SaveToWriter(writer io.Writer) error {
	_, err := e.File.WriteTo(writer)
	return err
}

func (e *SpreadsheetExcelize) Close() error {
	return e.File.Close()
}

func (e *SpreadsheetExcelize) GetSheetName() string {
	return e.SheetName
}

func (e *SpreadsheetExcelize) SetSheetName(name string) {
	e.SheetName = name
}

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

func (e *SpreadsheetExcelize) SetActiveSheet() error {
	index, err := e.File.GetSheetIndex(e.SheetName)
	if err != nil {
		return fmt.Errorf("failed to get sheet index: %w", err)
	}
	e.File.SetActiveSheet(index)
	return nil
}

func (e *SpreadsheetExcelize) SetColumnWidth(colLetter string, width float64) error {
	return e.File.SetColWidth(e.SheetName, colLetter, colLetter, width)
}

// Delegation to Excelize table operations
// These methods delegate to the TableExcelize adapter for cell and range operations.

func (e *SpreadsheetExcelize) GetCellValue(col, row int) (string, error) {
	return e.Table.GetCellValue(col, row)
}

func (e *SpreadsheetExcelize) SetCellValue(col, row int, value interface{}) error {
	return e.Table.SetCellValue(col, row, value)
}

func (e *SpreadsheetExcelize) MergeCells(startCol, startRow, endCol, endRow int) error {
	return e.Table.MergeCells(startCol, startRow, endCol, endRow)
}

func (e *SpreadsheetExcelize) IsCellMerged(col, row int) bool {
	return e.Table.IsCellMerged(col, row)
}

func (e *SpreadsheetExcelize) IsCellMergedHorizontally(col, row int) bool {
	return e.Table.IsCellMergedHorizontally(col, row)
}

func (e *SpreadsheetExcelize) ApplyBorderToCell(col, row int, side string, border *Border) error {
	return e.Table.ApplyBorderToCell(col, row, side, border)
}

func (e *SpreadsheetExcelize) ApplyBordersToRange(startCol, startRow, endCol, endRow int, borders Borders) error {
	return e.Table.ApplyBordersToRange(startCol, startRow, endCol, endRow, borders)
}

func (e *SpreadsheetExcelize) HasExistingBorder(col, row int, side string) bool {
	return e.Table.HasExistingBorder(col, row, side)
}

func (e *SpreadsheetExcelize) ApplyStyleToCell(col, row int, style Style) error {
	return e.Table.ApplyStyleToCell(col, row, style)
}

func (e *SpreadsheetExcelize) ApplyStyleToRange(startCol, startRow, endCol, endRow int, style Style) error {
	return e.Table.ApplyStyleToRange(startCol, startRow, endCol, endRow, style)
}

func (e *SpreadsheetExcelize) GetColumnLetter(col int) string {
	return e.Table.GetColumnLetter(col)
}

func (e *SpreadsheetExcelize) ProcessValue(value interface{}, format string) (interface{}, error) {
	return e.Table.ProcessValue(value, format)
}
