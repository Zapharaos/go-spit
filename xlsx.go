// xlsx.go - XLSX export logic for go-spit
//
// This file provides functions to write tabular data to XLSX files using a dynamic spreadsheet implementation.
// Supports multiple spreadsheet backends (e.g., Excelize) and advanced table features.

package go_spit

import (
	"fmt"
	"io"

	"github.com/Zapharaos/go-spit/internal/logger"
	"github.com/xuri/excelize/v2"
)

// XLSX represents the XLSX format implementation with dynamic spreadsheet implementation
type XLSX struct {
	Spreadsheet Spreadsheet
}

// NewXlsx creates a new XLSX instance with the provided Spreadsheet implementation
func NewXlsx(s Spreadsheet) *XLSX {
	return &XLSX{
		Spreadsheet: s,
	}
}

// NewXlsxWithExcelize creates a new XLSX instance with an TableExcelize implementation
func NewXlsxWithExcelize(file *excelize.File, sheetName string, t *Table) *XLSX {
	return &XLSX{
		Spreadsheet: NewSpreadsheetExcelize(file, sheetName, t),
	}
}

// WriteDataToFile writes data to file using the generic file writer
func (xlsx *XLSX) WriteDataToFile(options FileWriteOptions) (*FileWriteResult, error) {
	// Ensure extension is set for XLSX files
	if options.extension == "" {
		options.extension = FormatXSLX.String()
	}

	writeFunc := func(writer io.Writer) error {
		// Create new Excel file using the dynamic operations
		if err := xlsx.Spreadsheet.createNewFile(); err != nil {
			return fmt.Errorf("failed to create new Excel file: %w", err)
		}

		defer func() {
			_ = xlsx.Spreadsheet.close()
		}()

		// Write data to the file
		if err := xlsx.writeData(); err != nil {
			logger.L().Warn("Failed to write data to Excel file", logger.Error(err))
		}

		// Write to the writer using the dynamic operations
		if err := xlsx.Spreadsheet.saveToWriter(writer); err != nil {
			return fmt.Errorf("failed to write XLSX to writer: %w", err)
		}

		return nil
	}

	return options.writeToFile(writeFunc)
}

// writeData writes the provided data to the XLSX file
func (xlsx *XLSX) writeData() error {
	if xlsx.Spreadsheet.getSheetName() == "" {
		xlsx.Spreadsheet.setSheetName("Sheet1")
	}

	// Create sheet using dynamic operations
	if err := xlsx.Spreadsheet.createSheet(); err != nil {
		return fmt.Errorf("failed to create sheet: %w", err)
	}

	if err := xlsx.Spreadsheet.setActiveSheet(); err != nil {
		return fmt.Errorf("failed to set active sheet: %w", err)
	}

	// Get the table from the spreadsheet
	t := xlsx.Spreadsheet.getTable()
	if t == nil {
		return fmt.Errorf("no table data provided")
	}

	// Write headers if requested
	currentRow := 1
	if len(t.Columns) > 0 {
		headerRows := xlsx.writeHeaders()
		currentRow += headerRows
	}

	// Write data rows
	for _, item := range t.Data {
		colIndex := 1
		flatColumns := t.Columns.getFlattenedColumns()
		for _, column := range flatColumns {
			err := xlsx.writeCell(item, column, colIndex, currentRow)
			if err != nil {
				return err
			}
			colIndex++
		}
		currentRow++
	}

	// Render additional data like merging and styles
	if err := t.processMerging(xlsx.Spreadsheet); err != nil {
		logger.L().Warn("Failed to render merging data to Excel file", logger.Error(err))
	}

	if err := t.renderStyles(xlsx.Spreadsheet); err != nil {
		logger.L().Warn("Failed to render styles to Excel file", logger.Error(err))
	}

	// Auto-fit columns
	if err := xlsx.autoFitColumns(); err != nil {
		logger.L().Warn("Failed to auto-fit columns", logger.Error(err))
	}

	return nil
}

// writeHeaders writes multi-level headers to the Excel sheet
func (xlsx *XLSX) writeHeaders() int {
	// Get the table from the spreadsheet
	t := xlsx.Spreadsheet.getTable()
	nbColumns := len(t.Columns)

	// If no columns are defined, return 0
	if nbColumns == 0 {
		return 0
	}

	maxDepth := t.Columns.getMaxDepth()
	if maxDepth == 1 {
		// Simple single-level headers
		for i, column := range t.Columns {
			if err := xlsx.Spreadsheet.setCellValue(i+1, 1, column.Label); err != nil {
				logger.L().Warn("Failed to set header cell value", logger.Error(err))
			}
		}
		return 1
	}

	xlsx.writeHeaderRow(t.Columns, 1, maxDepth, 1)
	return maxDepth
}

// writeHeaderRow writes a specific header row, handling hierarchical structure
func (xlsx *XLSX) writeHeaderRow(columns Columns, currentRow, maxDepth, startCol int) int {
	currentCol := startCol

	for _, column := range columns {
		// Write the header cell value
		if err := xlsx.Spreadsheet.setCellValue(currentCol, currentRow, column.Label); err != nil {
			logger.L().Warn("Failed to set header cell value", logger.Error(err))
		}

		if column.hasSubColumns() {
			// Process sub-columns recursively for hierarchical headers
			if currentRow < maxDepth {
				xlsx.writeHeaderRow(column.Columns, currentRow+1, maxDepth, currentCol)
			}
			// Move to next column position after all sub-columns
			currentCol += column.getColumnCount()
		} else {
			// Simple leaf column - move to next position
			currentCol++
		}
	}

	return currentCol
}

// writeCell writes a single cell item
func (xlsx *XLSX) writeCell(item Data, column Column, colIndex, rowIndex int) error {
	value, err := item.lookup(column.Name)
	if err != nil {
		return nil // Skip missing values
	}

	processedValue, err := xlsx.Spreadsheet.processValue(value, column.Format)
	if err != nil {
		return fmt.Errorf("error processing value for column %s: %w", column.Name, err)
	}

	if err = xlsx.Spreadsheet.setCellValue(colIndex, rowIndex, processedValue); err != nil {
		return fmt.Errorf("failed to set cell value: %w", err)
	}

	return nil
}

// autoFitColumns auto-fits column widths using dynamic operations
func (xlsx *XLSX) autoFitColumns() error {
	for i := 1; i <= len(xlsx.Spreadsheet.getTable().Columns.getFlattenedColumns()); i++ {
		colLetter := xlsx.Spreadsheet.getColumnLetter(i)
		if err := xlsx.Spreadsheet.setColumnWidth(colLetter, 15); err != nil {
			return err
		}
	}
	return nil
}
