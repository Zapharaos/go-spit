package spreadsheet

import (
	"fmt"
	"io"

	"github.com/Zapharaos/go-spit/internal/table"
	"github.com/xuri/excelize/v2"
)

// Excelize provides Excelize-specific operations for spreadsheet handling
type Excelize struct {
	File      *excelize.File
	SheetName string
	Table     *table.Excelize
}

func NewExcelize(file *excelize.File, sheetName string, t *table.Table) *Excelize {
	return &Excelize{
		File:      file,
		SheetName: sheetName,
		Table:     table.NewExcelize(file, sheetName, t),
	}
}

func (e *Excelize) GetTable() *table.Table {
	return e.Table.GetTable()
}

func (e *Excelize) GetFile() interface{} {
	return e.File
}

func (e *Excelize) CreateNewFile() error {
	e.File = excelize.NewFile()
	return nil
}

func (e *Excelize) SaveToWriter(writer io.Writer) error {
	_, err := e.File.WriteTo(writer)
	return err
}

func (e *Excelize) Close() error {
	return e.File.Close()
}

func (e *Excelize) GetSheetName() string {
	return e.SheetName
}

func (e *Excelize) SetSheetName(name string) {
	e.SheetName = name
}

func (e *Excelize) CreateSheet() error {
	index, err := e.File.GetSheetIndex(e.SheetName)
	if err != nil || index == -1 {
		index, err = e.File.NewSheet(e.SheetName)
		if err != nil {
			return fmt.Errorf("failed to create sheet: %w", err)
		}
	}
	return nil
}

func (e *Excelize) SetActiveSheet() error {
	index, err := e.File.GetSheetIndex(e.SheetName)
	if err != nil {
		return fmt.Errorf("failed to get sheet index: %w", err)
	}
	e.File.SetActiveSheet(index)
	return nil
}

func (e *Excelize) SetColumnWidth(colLetter string, width float64) error {
	return e.File.SetColWidth(e.SheetName, colLetter, colLetter, width)
}

// Delegation to excelize table operations

func (e *Excelize) GetCellValue(col, row int) (string, error) {
	return e.Table.GetCellValue(col, row)
}

func (e *Excelize) SetCellValue(col, row int, value interface{}) error {
	return e.Table.SetCellValue(col, row, value)
}

func (e *Excelize) MergeCells(startCol, startRow, endCol, endRow int) error {
	return e.Table.MergeCells(startCol, startRow, endCol, endRow)
}

func (e *Excelize) IsCellMerged(col, row int) bool {
	return e.Table.IsCellMerged(col, row)
}

func (e *Excelize) IsCellMergedHorizontally(col, row int) bool {
	return e.Table.IsCellMergedHorizontally(col, row)
}

func (e *Excelize) ApplyCellBorder(col, row int, side string, borderSide *table.BorderSide) error {
	return e.Table.ApplyCellBorder(col, row, side, borderSide)
}

func (e *Excelize) ApplyRangeBorder(startCol, startRow, endCol, endRow int, borderConfig table.BorderConfig) error {
	return e.Table.ApplyRangeBorder(startCol, startRow, endCol, endRow, borderConfig)
}

func (e *Excelize) HasExistingBorder(col, row int, side string) bool {
	return e.Table.HasExistingBorder(col, row, side)
}

func (e *Excelize) ApplyCellStyle(col, row int, style table.StyleConfig) error {
	return e.Table.ApplyCellStyle(col, row, style)
}

func (e *Excelize) ApplyRangeStyle(startCol, startRow, endCol, endRow int, style table.StyleConfig) error {
	return e.Table.ApplyRangeStyle(startCol, startRow, endCol, endRow, style)
}

func (e *Excelize) GetColumnLetter(col int) string {
	return e.Table.GetColumnLetter(col)
}

func (e *Excelize) ProcessValue(value interface{}, format string) (interface{}, error) {
	return e.Table.ProcessValue(value, format)
}
