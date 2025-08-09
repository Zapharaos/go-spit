package go_spit

import (
	"fmt"
	"io"

	"github.com/Zapharaos/go-spit/internal/file"
	"github.com/Zapharaos/go-spit/internal/logger"
	"github.com/Zapharaos/go-spit/internal/spreadsheet"
	"github.com/Zapharaos/go-spit/internal/table"
	"github.com/xuri/excelize/v2"
)

// XLSX represents the XLSX format implementation with dynamic spreadsheet implementation
type XLSX struct {
	Spreadsheet spreadsheet.Spreadsheet
}

// NewXlsx creates a new XLSX instance with the provided Spreadsheet implementation
func NewXlsx(s spreadsheet.Spreadsheet) *XLSX {
	return &XLSX{
		Spreadsheet: s,
	}
}

// NewXlsxWithExcelize creates a new XLSX instance with an Excelize implementation
func NewXlsxWithExcelize(file *excelize.File, sheetName string, t *table.Table) *XLSX {
	return &XLSX{
		Spreadsheet: spreadsheet.NewExcelize(file, sheetName, t),
	}
}

// WriteDataToFile writes data to file using the generic file writer
func (xlsx *XLSX) WriteDataToFile(options file.WriteOptions) (*file.WriteResult, error) {
	// Ensure extension is set for XLSX files
	if options.Extension == "" {
		options.Extension = FormatXSLX.String()
	}

	writeFunc := func(writer io.Writer) error {
		// Create new Excel file using the dynamic operations
		if err := xlsx.Spreadsheet.CreateNewFile(); err != nil {
			return fmt.Errorf("failed to create new Excel file: %w", err)
		}

		defer func() {
			_ = xlsx.Spreadsheet.Close()
		}()

		// Write data to the file
		if err := xlsx.writeData(); err != nil {
			logger.L().Warn("Failed to write data to Excel file", logger.Error(err))
		}

		// Write to the writer using the dynamic operations
		if err := xlsx.Spreadsheet.SaveToWriter(writer); err != nil {
			return fmt.Errorf("failed to write XLSX to writer: %w", err)
		}

		return nil
	}

	return options.WriteToFile(writeFunc)
}

// writeData writes the provided data to the XLSX file
func (xlsx *XLSX) writeData() error {
	if xlsx.Spreadsheet.GetSheetName() == "" {
		xlsx.Spreadsheet.SetSheetName("Sheet1")
	}

	// Create sheet using dynamic operations
	if err := xlsx.Spreadsheet.CreateSheet(); err != nil {
		return fmt.Errorf("failed to create sheet: %w", err)
	}

	if err := xlsx.Spreadsheet.SetActiveSheet(); err != nil {
		return fmt.Errorf("failed to set active sheet: %w", err)
	}

	// Get the table from the spreadsheet
	t := xlsx.Spreadsheet.GetTable()
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
		flatColumns := t.Columns.GetFlattenedColumns()
		for _, column := range flatColumns {
			err := xlsx.writeCell(item, column, colIndex, currentRow)
			if err != nil {
				return err
			}
			colIndex++
		}
		currentRow++
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
	t := xlsx.Spreadsheet.GetTable()
	nbColumns := len(t.Columns)

	// If no columns are defined, return 0
	if nbColumns == 0 {
		return 0
	}

	maxDepth := t.Columns.GetMaxDepth()
	if maxDepth == 1 {
		// Simple single-level headers
		for i, column := range t.Columns {
			if err := xlsx.Spreadsheet.SetCellValue(i+1, 1, column.Label); err != nil {
				logger.L().Warn("Failed to set header cell value", logger.Error(err))
			}
		}

		if err := xlsx.styleHeaders(nbColumns, 1); err != nil {
			logger.L().Warn("Failed to apply header styling", logger.Error(err))
		}
		return 1
	}

	xlsx.writeHeaderRow(t.Columns, 1, maxDepth, 1)

	totalColumns := t.Columns.GetTotalColumnCount()
	if err := xlsx.styleHeaders(totalColumns, maxDepth); err != nil {
		logger.L().Warn("Failed to apply header styling", logger.Error(err))
	}

	return maxDepth
}

// writeHeaderRow writes a specific header row, handling merging and alignment
func (xlsx *XLSX) writeHeaderRow(columns table.Columns, currentRow, maxDepth, startCol int) int {
	currentCol := startCol

	for _, column := range columns {
		if err := xlsx.Spreadsheet.SetCellValue(currentCol, currentRow, column.Label); err != nil {
			logger.L().Warn("Failed to set header cell value", logger.Error(err))
		}

		if column.HasSubColumns() {
			columnSpan := column.GetColumnCount()
			endCol := currentCol + columnSpan - 1
			if endCol > currentCol {
				if err := xlsx.Spreadsheet.MergeCells(currentCol, currentRow, endCol, currentRow); err != nil {
					logger.L().Warn("Failed to merge header cells horizontally", logger.Error(err))
				}
			}

			if currentRow < maxDepth {
				xlsx.writeHeaderRow(column.Columns, currentRow+1, maxDepth, currentCol)
			}
			currentCol += columnSpan
		} else {
			if currentRow < maxDepth {
				if err := xlsx.Spreadsheet.MergeCells(currentCol, currentRow, currentCol, maxDepth); err != nil {
					logger.L().Warn("Failed to merge header cells vertically", logger.Error(err))
				}
			}
			currentCol++
		}
	}

	return currentCol
}

// styleHeaders applies styling to multi-level headers
func (xlsx *XLSX) styleHeaders(col, rows int) error {
	headerStyle := table.StyleConfig{
		Bold:            true,
		BackgroundColor: "#E0E0E0",
		Alignment:       table.AlignmentCenter,
	}

	return xlsx.Spreadsheet.ApplyRangeStyle(1, 1, col, rows, headerStyle)
}

// writeCell writes a single cell item
func (xlsx *XLSX) writeCell(item table.Data, column table.Column, colIndex, rowIndex int) error {
	value, err := item.Lookup(column.Name)
	if err != nil {
		return nil // Skip missing values
	}

	processedValue, err := xlsx.Spreadsheet.ProcessValue(value, column.Format)
	if err != nil {
		return fmt.Errorf("error processing value for column %s: %w", column.Name, err)
	}

	if err = xlsx.Spreadsheet.SetCellValue(colIndex, rowIndex, processedValue); err != nil {
		return fmt.Errorf("failed to set cell value: %w", err)
	}

	return nil
}

// autoFitColumns auto-fits column widths using dynamic operations
func (xlsx *XLSX) autoFitColumns() error {
	for i := 1; i <= len(xlsx.Spreadsheet.GetTable().Columns.GetFlattenedColumns()); i++ {
		colLetter := xlsx.Spreadsheet.GetColumnLetter(i)
		if err := xlsx.Spreadsheet.SetColumnWidth(colLetter, 15); err != nil {
			return err
		}
	}
	return nil
}
