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
	return ExportXLSXSheets([]Spreadsheet{s}, params)
}

// ExportXLSXSheets writes data for one or more sheets to a single XLSX file.
// Each Spreadsheet in the slice represents one sheet; all sheets are written to the same underlying file.
// When the first sheet has no file, a new file is created and shared with the remaining sheets.
func ExportXLSXSheets(sheets []Spreadsheet, params FileWriteParams) (*FileWriteResult, error) {
	if len(sheets) == 0 {
		return nil, fmt.Errorf("no sheets provided")
	}

	// Ensure Extension is set for XLSX files
	if params.Extension == "" {
		params.Extension = FormatXSLX.String()
	}

	firstSheet := sheets[0]

	// Ensure the spreadsheet file is initialized
	f := firstSheet.GetFile()
	if f == nil || reflect.ValueOf(f).IsNil() {
		L().Debug("No existing spreadsheet file found, creating new one")
		if err := firstSheet.CreateNewFile(); err != nil {
			L().Error("Failed to create new XLSX file", Error(err))
			return nil, fmt.Errorf("failed to create new XLSX file: %w", err)
		}

		defer func() {
			if err := firstSheet.Close(); err != nil {
				L().Warn("Error closing spreadsheet", Error(err))
			}
		}()
	}

	// Propagate the file to all other sheets that do not already have one.
	// GetFile is called again here only when there are multiple sheets to initialise.
	if len(sheets) > 1 {
		f = firstSheet.GetFile()
		for _, sheet := range sheets[1:] {
			sheetF := sheet.GetFile()
			if sheetF == nil || reflect.ValueOf(sheetF).IsNil() {
				if err := sheet.InitWithFile(f); err != nil {
					L().Error("Failed to initialize sheet with existing file", Error(err))
					return nil, fmt.Errorf("failed to initialize sheet with existing file: %w", err)
				}
			}
		}
	}

	L().Info("Starting XLSX export to file", String("filename", params.Filename))

	// Create a write function that handles the XLSX file creation and writing
	writeFunc := func(writer io.Writer) error {
		for _, sheet := range sheets {
			xlsxConfig := &xlsx{
				spreadsheet: sheet,
				params:      params,
			}

			L().Debug("Writing data to sheet")
			if err := xlsxConfig.writeData(); err != nil {
				return fmt.Errorf("failed to write data to XLSX file: %w", err)
			}
		}

		L().Debug("Saving Excel file to writer")
		if err := firstSheet.SaveToWriter(writer); err != nil {
			return fmt.Errorf("failed to write XLSX to writer: %w", err)
		}

		return nil
	}

	// Use the generic file writer to handle the actual file writing
	result, err := params.WriteToFile(writeFunc)
	if err != nil {
		L().Error("Failed to write XLSX to file", Error(err))
		return nil, err
	}

	L().Info("XLSX export completed", String("filename", params.Filename))
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
	if xlsx.spreadsheet.GetSheetName() == "" {
		xlsx.spreadsheet.SetSheetName("Sheet1")
	}

	L().Debug("Creating sheet")
	if err := xlsx.spreadsheet.CreateSheet(); err != nil {
		return fmt.Errorf("failed to create sheet: %w", err)
	}

	if err := xlsx.spreadsheet.SetActiveSheet(); err != nil {
		return fmt.Errorf("failed to set active sheet: %w", err)
	}

	t := xlsx.spreadsheet.GetTable()
	if t == nil {
		return fmt.Errorf("no table data provided")
	}

	currentRow := 1
	if len(t.Preamble) > 0 {
		L().Debug("Writing preamble rows")
		preambleRows, err := xlsx.writePreamble(currentRow)
		if err != nil {
			return fmt.Errorf("failed to write preamble: %w", err)
		}
		currentRow += preambleRows
	}

	if len(t.Columns) > 0 {
		L().Debug("Writing headers")
		headerRows, err := xlsx.writeHeaders(currentRow)
		if err != nil {
			return fmt.Errorf("failed to write headers: %w", err)
		}
		currentRow += headerRows
	}

	L().Debug("Writing data rows")
	flatColumns := t.Columns.GetFlattenedColumns()
	for _, item := range t.Data {
		colIndex := 1
		for _, column := range flatColumns {
			if err := xlsx.writeCell(item, column, colIndex, currentRow); err != nil {
				return fmt.Errorf("failed to write cell: %w", err)
			}
			colIndex++
		}
		currentRow++
	}

	xlsx.autoFitColumns()

	if err := t.ProcessMerging(xlsx.spreadsheet); err != nil {
		return fmt.Errorf("failed to process merging: %w", err)
	}

	if err := t.RenderStyles(xlsx.spreadsheet); err != nil {
		return fmt.Errorf("failed to render styles: %w", err)
	}

	L().Debug("XLSX data writing complete.")
	return nil
}

// writeHeaders writes multi-level headers to the Excel sheet starting at the given row.
// Returns the number of header rows written and error if any header cell fails to write.
func (xlsx *xlsx) writeHeaders(startRow int) (int, error) {
	t := xlsx.spreadsheet.GetTable()
	nbColumns := len(t.Columns)

	if nbColumns == 0 {
		L().Warn("No columns defined for headers")
		return 0, nil
	}

	maxDepth := t.Columns.GetMaxDepth()
	if maxDepth == 1 {
		for i, column := range t.Columns {
			if err := xlsx.spreadsheet.SetCellValue(i+1, startRow, column.Label); err != nil {
				return 0, fmt.Errorf("failed to set header cell value for column %s: %w", column.Name, err)
			}
		}
		return 1, nil
	}

	L().Debug("Writing multi-level headers", Int("maxDepth", maxDepth))
	maxRow := startRow + maxDepth - 1
	if err := xlsx.writeHeaderRow(t.Columns, startRow, maxRow, 1); err != nil {
		return 0, err
	}
	return maxDepth, nil
}

// writeHeaderRow writes a specific header row, handling hierarchical structure.
// Recursively processes sub-columns for multi-level headers.
func (xlsx *xlsx) writeHeaderRow(columns Columns, currentRow, maxRow, startCol int) error {
	currentCol := startCol

	for _, column := range columns {
		if err := xlsx.spreadsheet.SetCellValue(currentCol, currentRow, column.Label); err != nil {
			return fmt.Errorf("failed to set header cell value for column %s at (%d, %d): %w", column.Name, currentCol, currentRow, err)
		}

		if column.HasSubColumns() {
			// Process sub-columns recursively for hierarchical headers
			if currentRow < maxRow {
				if err := xlsx.writeHeaderRow(column.Columns, currentRow+1, maxRow, currentCol); err != nil {
					return err
				}
			}
			// Move to next column position after all sub-columns
			currentCol += column.CountSubColumns()
		} else {
			// Simple leaf column - move to next position
			currentCol++
		}
	}

	return nil
}

// writePreamble writes free-form preamble rows to the sheet starting at startRow.
// Returns the number of rows written.
func (xlsx *xlsx) writePreamble(startRow int) (int, error) {
	t := xlsx.spreadsheet.GetTable()
	for i, row := range t.Preamble {
		for j, val := range row.Values {
			if err := xlsx.spreadsheet.SetCellValue(j+1, startRow+i, val); err != nil {
				return 0, fmt.Errorf("failed to write preamble cell at (%d, %d): %w", j+1, startRow+i, err)
			}
		}
	}
	return len(t.Preamble), nil
}

// writeCell writes a single cell item to the spreadsheet.
// Looks up the value, processes formatting, and sets the cell value.
// Special formats (formula, hyperlink, default) trigger dedicated Excelize operations.
func (xlsx *xlsx) writeCell(item Data, column *Column, colIndex, rowIndex int) error {
	value, err, found := item.Lookup(column.Name)
	if err == nil && !found {
		return nil
	}
	if err != nil {
		return fmt.Errorf("error looking up value for column %s: %w", column.Name, err)
	}

	// Image values are inserted as cell-anchored pictures rather than text.
	if img, ok := asImage(value); ok {
		if err = xlsx.spreadsheet.SetCellImage(colIndex, rowIndex, img); err != nil {
			return fmt.Errorf("error setting image for column %s at (%d, %d): %w", column.Name, colIndex, rowIndex, err)
		}
		return nil
	}

	processedValue, err := xlsx.spreadsheet.ProcessValue(value, column.Format)
	if err != nil {
		return fmt.Errorf("error processing value %s for column %s: %w", value, column.Name, err)
	}

	switch column.Format {
	case ExcelizeFormatFormula:
		formula := fmt.Sprintf("%v", processedValue)
		if err = xlsx.spreadsheet.SetCellFormula(colIndex, rowIndex, formula); err != nil {
			return fmt.Errorf("error setting formula for column %s at (%d, %d): %w", column.Name, colIndex, rowIndex, err)
		}
	case ExcelizeFormatHyperlink:
		link := fmt.Sprintf("%v", processedValue)
		if err = xlsx.spreadsheet.SetCellValue(colIndex, rowIndex, link); err != nil {
			return fmt.Errorf("error setting cell value for column %s at (%d, %d): %w", column.Name, colIndex, rowIndex, err)
		}
		if err = xlsx.spreadsheet.SetCellHyperLink(colIndex, rowIndex, link); err != nil {
			return fmt.Errorf("error setting hyperlink for column %s at (%d, %d): %w", column.Name, colIndex, rowIndex, err)
		}
	default:
		if err = xlsx.spreadsheet.SetCellValue(colIndex, rowIndex, processedValue); err != nil {
			return fmt.Errorf("error setting cell value for column %s at (%d, %d): %w", column.Name, colIndex, rowIndex, err)
		}
	}

	return nil
}

// autoFitColumns auto-fits column widths using dynamic operations.
// Uses the column-specific width when set, otherwise falls back to a default width of 15.
func (xlsx *xlsx) autoFitColumns() {
	const defaultWidth = 15
	flatColumns := xlsx.spreadsheet.GetTable().Columns.GetFlattenedColumns()
	for i, column := range flatColumns {
		colLetter := xlsx.spreadsheet.GetColumnLetter(i + 1)
		width := column.Width
		if width <= 0 {
			width = defaultWidth
		}
		if err := xlsx.spreadsheet.SetColumnWidth(colLetter, width); err != nil {
			L().Warn("Failed to set column width", String("column", colLetter), Error(err))
		}
	}
}
