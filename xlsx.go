// xlsx.go - XLSX export logic for go-spit
//
// This file provides functions to write tabular data to XLSX files using a dynamic spreadsheet implementation.
// Supports multiple spreadsheet backends (e.g., Excelize) and advanced table features.

package spit

import (
	"fmt"
	"io"
	"reflect"
)

// ExportXLSX writes table data to an XLSX file using the generic file writer and a dynamic spreadsheet implementation.
func ExportXLSX(s Spreadsheet, params FileWriteParams) (*FileWriteResult, error) {
	// Ensure extension is set for XLSX files
	if params.extension == "" {
		params.extension = FormatXSLX.String()
	}

	xlsxConfig := &xlsx{
		spreadsheet: s,
		params:      params,
	}

	// Ensure the spreadsheet file is initialized
	f := xlsxConfig.spreadsheet.getFile()
	if f == nil || reflect.ValueOf(f).IsNil() {
		L().Debug("No existing spreadsheet file found, creating new one")
		if err := xlsxConfig.spreadsheet.createNewFile(); err != nil {
			L().Error("Failed to create new XLSX file", Error(err))
			return nil, fmt.Errorf("failed to create new XLSX file: %w", err)
		}

		defer func() {
			if err := xlsxConfig.spreadsheet.close(); err != nil {
				L().Warn("Error closing spreadsheet", Error(err))
			}
		}()
	}

	L().Info("Starting XLSX export to file", String("filename", xlsxConfig.params.Filename))

	// Create a write function that handles the XLSX file creation and writing
	writeFunc := func(writer io.Writer) error {

		L().Debug("Writing data to Excel file")
		if err := xlsxConfig.writeData(); err != nil {
			return fmt.Errorf("failed to write data to XLSX file: %w", err)
		}

		L().Debug("Saving Excel file to writer")
		if err := xlsxConfig.spreadsheet.saveToWriter(writer); err != nil {
			return fmt.Errorf("failed to write XLSX to writer: %w", err)
		}

		return nil
	}

	// Use the generic file writer to handle the actual file writing
	result, err := xlsxConfig.params.writeToFile(writeFunc)
	if err != nil {
		L().Error("Failed to write XLSX to file", Error(err))
		return nil, err
	}

	L().Info("XLSX export completed", String("filename", xlsxConfig.params.Filename))
	return result, nil
}

// xlsx represents the XLSX format implementation with dynamic spreadsheet implementation
type xlsx struct {
	spreadsheet Spreadsheet
	params      FileWriteParams
}

// writeData writes the provided table data to the XLSX file.
// Handles sheet creation, header writing, data rows, merging, styling, and auto-fitting columns.
func (xlsx *xlsx) writeData() error {
	if xlsx.spreadsheet.getSheetName() == "" {
		xlsx.spreadsheet.setSheetName("Sheet1")
	}

	L().Debug("Creating sheet")
	if err := xlsx.spreadsheet.createSheet(); err != nil {
		return fmt.Errorf("failed to create sheet: %w", err)
	}

	if err := xlsx.spreadsheet.setActiveSheet(); err != nil {
		return fmt.Errorf("failed to set active sheet: %w", err)
	}

	t := xlsx.spreadsheet.getTable()
	if t == nil {
		return fmt.Errorf("no table data provided")
	}

	currentRow := 1
	if len(t.Columns) > 0 {
		L().Debug("Writing headers")
		headerRows, err := xlsx.writeHeaders()
		if err != nil {
			return fmt.Errorf("failed to write headers: %w", err)
		}
		currentRow += headerRows
	}

	L().Debug("Writing data rows")
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

	if err := t.processMerging(xlsx.spreadsheet); err != nil {
		return fmt.Errorf("failed to process merging: %w", err)
	}

	if err := t.renderStyles(xlsx.spreadsheet); err != nil {
		return fmt.Errorf("failed to render styles: %w", err)
	}

	L().Debug("XLSX data writing complete.")
	return nil
}

// writeHeaders writes multi-level headers to the Excel sheet.
// Returns the number of header rows written and error if any header cell fails to write.
func (xlsx *xlsx) writeHeaders() (int, error) {
	t := xlsx.spreadsheet.getTable()
	nbColumns := len(t.Columns)

	if nbColumns == 0 {
		L().Warn("No columns defined for headers")
		return 0, nil
	}

	maxDepth := t.Columns.getMaxDepth()
	if maxDepth == 1 {
		for i, column := range t.Columns {
			if err := xlsx.spreadsheet.setCellValue(i+1, 1, column.Label); err != nil {
				return 0, fmt.Errorf("failed to set header cell value for column %s: %w", column.Name, err)
			}
		}
		return 1, nil
	}

	L().Debug("Writing multi-level headers", Int("maxDepth", maxDepth))
	if err := xlsx.writeHeaderRow(t.Columns, 1, maxDepth, 1); err != nil {
		return 0, err
	}
	return maxDepth, nil
}

// writeHeaderRow writes a specific header row, handling hierarchical structure.
// Recursively processes sub-columns for multi-level headers.
func (xlsx *xlsx) writeHeaderRow(columns Columns, currentRow, maxDepth, startCol int) error {
	currentCol := startCol

	for _, column := range columns {
		if err := xlsx.spreadsheet.setCellValue(currentCol, currentRow, column.Label); err != nil {
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

// writeCell writes a single cell item to the spreadsheet.
// Looks up the value, processes formatting, and sets the cell value.
func (xlsx *xlsx) writeCell(item Data, column Column, colIndex, rowIndex int) error {
	value, err := item.lookup(column.Name)
	if err != nil {
		return fmt.Errorf("error looking up value for column %s: %w", column.Name, err)
	}

	processedValue, err := xlsx.spreadsheet.processValue(value, column.Format)
	if err != nil {
		return fmt.Errorf("error processing value %s for column %s: %w", value, column.Name, err)
	}

	if err = xlsx.spreadsheet.setCellValue(colIndex, rowIndex, processedValue); err != nil {
		return fmt.Errorf("error setting cell value for column %s at (%d, %d): %w", column.Name, colIndex, rowIndex, err)
	}

	return nil
}

// autoFitColumns auto-fits column widths using dynamic operations.
// Sets a default width for each column for improved readability.
func (xlsx *xlsx) autoFitColumns() {
	for i := 1; i <= len(xlsx.spreadsheet.getTable().Columns.getFlattenedColumns()); i++ {
		colLetter := xlsx.spreadsheet.getColumnLetter(i)
		if err := xlsx.spreadsheet.setColumnWidth(colLetter, 15); err != nil {
			L().Warn("Failed to set column width", String("column", colLetter), Error(err))
		}
	}
}
