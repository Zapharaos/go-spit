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
	logger.L().Info("Starting XLSX export to file", logger.String("filename", options.Filename))

	// Ensure extension is set for XLSX files
	if options.extension == "" {
		options.extension = FormatXSLX.String()
	}

	// Create a write function that handles the XLSX file creation and writing
	writeFunc := func(writer io.Writer) error {
		logger.L().Debug("Creating new Excel file")
		if err := xlsx.Spreadsheet.createNewFile(); err != nil {
			return fmt.Errorf("failed to create new XLSX file: %w", err)
		}

		defer func() {
			if err := xlsx.Spreadsheet.close(); err != nil {
				logger.L().Warn("Error closing spreadsheet", logger.Error(err))
			}
		}()

		logger.L().Debug("Writing data to Excel file")
		if err := xlsx.writeData(); err != nil {
			return fmt.Errorf("failed to write data to XLSX file: %w", err)
		}

		logger.L().Debug("Saving Excel file to writer")
		if err := xlsx.Spreadsheet.saveToWriter(writer); err != nil {
			return fmt.Errorf("failed to write XLSX to writer: %w", err)
		}

		return nil
	}

	// Use the generic file writer to handle the actual file writing
	result, err := options.writeToFile(writeFunc)
	if err != nil {
		logger.L().Error("Failed to write XLSX to file", logger.Error(err))
		return nil, err
	}

	logger.L().Info("XLSX export completed", logger.String("filename", options.Filename))
	return result, nil
}

// writeData writes the provided data to the XLSX file
func (xlsx *XLSX) writeData() error {
	if xlsx.Spreadsheet.getSheetName() == "" {
		xlsx.Spreadsheet.setSheetName("Sheet1")
	}

	logger.L().Debug("Creating sheet")
	if err := xlsx.Spreadsheet.createSheet(); err != nil {
		return fmt.Errorf("failed to create sheet: %w", err)
	}

	if err := xlsx.Spreadsheet.setActiveSheet(); err != nil {
		return fmt.Errorf("failed to set active sheet: %w", err)
	}

	t := xlsx.Spreadsheet.getTable()
	if t == nil {
		return fmt.Errorf("no table data provided")
	}

	currentRow := 1
	if len(t.Columns) > 0 {
		logger.L().Debug("Writing headers")
		headerRows, err := xlsx.writeHeaders()
		if err != nil {
			return fmt.Errorf("failed to write headers: %w", err)
		}
		currentRow += headerRows
	}

	logger.L().Debug("Writing data rows")
	for _, item := range t.Data {
		colIndex := 1
		flatColumns := t.Columns.getFlattenedColumns()
		for _, column := range flatColumns {
			if err := xlsx.writeCell(item, column, colIndex, currentRow); err != nil {
				return fmt.Errorf("failed to write cell: %w", err)
			}
			colIndex++
		}
		currentRow++
	}

	xlsx.autoFitColumns()

	if err := t.processMerging(xlsx.Spreadsheet); err != nil {
		return fmt.Errorf("failed to process merging: %w", err)
	}

	if err := t.renderStyles(xlsx.Spreadsheet); err != nil {
		return fmt.Errorf("failed to render styles: %w", err)
	}

	logger.L().Debug("XLSX data writing complete.")
	return nil
}

// writeHeaders writes multi-level headers to the Excel sheet
// Returns error and fails fast if any header cell fails to write
func (xlsx *XLSX) writeHeaders() (int, error) {
	t := xlsx.Spreadsheet.getTable()
	nbColumns := len(t.Columns)

	if nbColumns == 0 {
		logger.L().Warn("No columns defined for headers")
		return 0, nil
	}

	maxDepth := t.Columns.getMaxDepth()
	if maxDepth == 1 {
		for i, column := range t.Columns {
			if err := xlsx.Spreadsheet.setCellValue(i+1, 1, column.Label); err != nil {
				return 0, fmt.Errorf("failed to set header cell value for column %s: %w", column.Name, err)
			}
		}
		return 1, nil
	}

	logger.L().Debug("Writing multi-level headers", logger.Int("maxDepth", maxDepth))
	if err := xlsx.writeHeaderRow(t.Columns, 1, maxDepth, 1); err != nil {
		return 0, err
	}
	return maxDepth, nil
}

// writeHeaderRow writes a specific header row, handling hierarchical structure
func (xlsx *XLSX) writeHeaderRow(columns Columns, currentRow, maxDepth, startCol int) error {
	currentCol := startCol

	for _, column := range columns {
		if err := xlsx.Spreadsheet.setCellValue(currentCol, currentRow, column.Label); err != nil {
			return fmt.Errorf("failed to set header cell value for column %s at (%d, %d): %w", column.Name, currentCol, currentRow, err)
		}

		if column.hasSubColumns() {
			// Process sub-columns recursively for hierarchical headers
			if currentRow < maxDepth {
				if err := xlsx.writeHeaderRow(column.Columns, currentRow+1, maxDepth, currentCol); err != nil {
					return err
				}
			}
			// Move to next column position after all sub-columns
			currentCol += column.getColumnCount()
		} else {
			// Simple leaf column - move to next position
			currentCol++
		}
	}

	return nil
}

// writeCell writes a single cell item
func (xlsx *XLSX) writeCell(item Data, column Column, colIndex, rowIndex int) error {
	value, err := item.lookup(column.Name)
	if err != nil {
		return fmt.Errorf("error looking up value for column %s: %w", column.Name, err)
	}

	processedValue, err := xlsx.Spreadsheet.processValue(value, column.Format)
	if err != nil {
		return fmt.Errorf("error processing value %s for column %s: %w", value, column.Name, err)
	}

	if err = xlsx.Spreadsheet.setCellValue(colIndex, rowIndex, processedValue); err != nil {
		return fmt.Errorf("error setting cell value for column %s at (%d, %d): %w", column.Name, colIndex, rowIndex, err)
	}

	return nil
}

// autoFitColumns auto-fits column widths using dynamic operations
func (xlsx *XLSX) autoFitColumns() {
	for i := 1; i <= len(xlsx.Spreadsheet.getTable().Columns.getFlattenedColumns()); i++ {
		colLetter := xlsx.Spreadsheet.getColumnLetter(i)
		if err := xlsx.Spreadsheet.setColumnWidth(colLetter, 15); err != nil {
			logger.L().Warn("Failed to set column width", logger.String("column", colLetter), logger.Error(err))
		}
	}
}
