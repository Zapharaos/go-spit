package go_spit

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"io"
	"time"
)

// XLSX contains XLSX-specific export parameters
type XLSX struct {
	File          *excelize.File
	Data          DataSlice
	Columns       Columns
	RowConfigs    RowConfigs  // Optional row configurations
	CellConfigs   CellConfigs // Optional cell configurations
	Limit         int64
	ListSeparator string
	WriteHeader   bool
	SheetName     string
}

// GetColumnsLabel returns the label of the columns (XLSX-specific method)
func (xlsx XLSX) GetColumnsLabel() []string {
	return xlsx.Columns.GetColumnsLabels()
}

// WriteDataToFile writes data to file using the generic file writer
func (xlsx XLSX) WriteDataToFile(options FileWriteOptions) (*FileWriteResult, error) {
	// Ensure extension is set for XLSX files
	if options.Extension == "" {
		options.Extension = FormatXSLX.String()
	}

	writeFunc := func(writer io.Writer) error {
		// Create new Excel file
		xlsx.File = excelize.NewFile()
		defer func() {
			_ = xlsx.File.Close()
		}()

		// Write data to the file
		if err := xlsx.writeData(); err != nil {
			L().Warn("Failed to close Excel file", Error(err))
		}

		// Write to the writer
		if _, err := xlsx.File.WriteTo(writer); err != nil {
			return fmt.Errorf("failed to write XLSX to writer: %w", err)
		}

		return nil
	}

	return options.writeToFile(writeFunc)
}

// writeData writes the provided data to the XLSX file
func (xlsx XLSX) writeData() error {
	if xlsx.SheetName == "" {
		xlsx.SheetName = "Sheet1"
	}

	// Create or get the sheet
	index, err := xlsx.File.GetSheetIndex(xlsx.SheetName)
	if err != nil || index == -1 {
		index, err = xlsx.File.NewSheet(xlsx.SheetName)
		if err != nil {
			return fmt.Errorf("failed to create sheet: %w", err)
		}
	}

	xlsx.File.SetActiveSheet(index)

	currentRow := 1

	// Write headers if requested
	if xlsx.WriteHeader && len(xlsx.Columns) > 0 {
		headerRows := xlsx.writeHeaders()
		currentRow += headerRows
	}

	// Write data rows
	for _, item := range xlsx.Data {
		colIndex := 1
		// Use flattened columns to write data cells since data should only be written to leaf columns
		flatColumns := xlsx.Columns.GetFlattenedColumns()
		for _, column := range flatColumns {
			err = xlsx.writeCell(item, column, colIndex, currentRow)
			if err != nil {
				return err
			}
			colIndex++
		}
		currentRow++
	}

	return nil
}

// writeHeaders writes multi-level headers to the Excel sheet with support for arbitrary depth
func (xlsx XLSX) writeHeaders() int {
	if len(xlsx.Columns) == 0 {
		return 0
	}

	// Calculate the maximum depth to determine how many header rows we need
	maxDepth := xlsx.Columns.GetMaxDepth()
	if maxDepth == 1 {
		// Simple single-level headers
		for i, column := range xlsx.Columns {
			cellRef, err := excelize.CoordinatesToCellName(i+1, 1)
			if err != nil {
				L().Warn("Failed to get cell reference for header", Error(err))
				continue
			}

			if err = xlsx.File.SetCellValue(xlsx.SheetName, cellRef, column.Label); err != nil {
				L().Warn("Failed to set header cell value", Error(err))
			}
		}

		// Apply header styling
		if err := xlsx.styleHeaders(len(xlsx.Columns), 1); err != nil {
			L().Warn("Failed to apply header styling", Error(err))
		}
		return 1
	}

	xlsx.writeHeaderRow(xlsx.Columns, 1, maxDepth, 1)

	// Apply header styling to all rows
	totalColumns := xlsx.Columns.GetTotalColumnCount()
	if err := xlsx.styleHeaders(totalColumns, maxDepth); err != nil {
		L().Warn("Failed to apply header styling", Error(err))
	}

	return maxDepth
}

// writeHeaderRow writes a specific header row, handling merging and alignment
func (xlsx XLSX) writeHeaderRow(columns Columns, currentRow, maxDepth, startCol int) int {
	currentCol := startCol

	for _, column := range columns {
		cellRef, err := excelize.CoordinatesToCellName(currentCol, currentRow)
		if err != nil {
			L().Warn("Failed to get cell reference for header", Error(err))
			if column.HasSubColumns() {
				currentCol = xlsx.writeHeaderRow(column.Columns, currentRow, maxDepth, currentCol)
			} else {
				currentCol++
			}
			continue
		}

		if err = xlsx.File.SetCellValue(xlsx.SheetName, cellRef, column.Label); err != nil {
			L().Warn("Failed to set header cell value", Error(err))
		}

		if column.HasSubColumns() {
			// This column has sub-columns, so we need to merge horizontally across them
			columnSpan := column.GetColumnCount()
			endCol := currentCol + columnSpan - 1
			endCellRef, err := excelize.CoordinatesToCellName(endCol, currentRow)
			if err == nil && endCol > currentCol {
				if err = xlsx.File.MergeCell(xlsx.SheetName, cellRef, endCellRef); err != nil {
					L().Warn("Failed to merge header cells horizontally", Error(err))
				}
			}

			// Recursively write sub-columns for the next row
			if currentRow < maxDepth {
				xlsx.writeHeaderRow(column.Columns, currentRow+1, maxDepth, currentCol)
			}
			currentCol += columnSpan
		} else {
			// This is a leaf column, merge vertically down to the last header row
			if currentRow < maxDepth {
				endCellRef, err := excelize.CoordinatesToCellName(currentCol, maxDepth)
				if err == nil {
					if err = xlsx.File.MergeCell(xlsx.SheetName, cellRef, endCellRef); err != nil {
						L().Warn("Failed to merge header cells vertically", Error(err))
					}
				}
			}
			currentCol++
		}
	}

	return currentCol
}

// styleHeaders applies styling to multi-level headers
func (xlsx XLSX) styleHeaders(col, rows int) error {
	// Create header style
	styleID, err := xlsx.File.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold: true,
		},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#E0E0E0"},
			Pattern: 1,
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
	})
	if err != nil {
		return fmt.Errorf("failed to create header style: %w", err)
	}

	// Apply style to all header cells
	startCell, _ := excelize.CoordinatesToCellName(1, 1)
	endCell, _ := excelize.CoordinatesToCellName(col, rows)
	if err = xlsx.File.SetCellStyle(xlsx.SheetName, startCell, endCell, styleID); err != nil {
		return fmt.Errorf("failed to apply header style: %w", err)
	}

	return nil
}

// writeCell writes a single cell item
func (xlsx XLSX) writeCell(item Data, column Column, colIndex, rowIndex int) error {
	value, err := item.Lookup(column.Name)
	if err != nil {
		return nil // Skip missing values
	}

	// Process the value based on column format
	processedValue, err := xlsx.processValue(value, column.Format)
	if err != nil {
		return fmt.Errorf("error processing value for column %s: %w", column.Name)
	}

	cellRef, err := excelize.CoordinatesToCellName(colIndex, rowIndex)
	if err != nil {
		return fmt.Errorf("failed to get cell reference: %w", err)
	}

	if err = xlsx.File.SetCellValue(xlsx.SheetName, cellRef, processedValue); err != nil {
		return fmt.Errorf("failed to set cell value: %w", err)
	}

	return nil
}

// processValue processes a value for XLSX output based on its type and format
func (xlsx XLSX) processValue(value interface{}, format string) (interface{}, error) {
	switch v := value.(type) {
	case []interface{}:
		if xlsx.ListSeparator != "" {
			return convertSliceToString(v, format, xlsx.ListSeparator)
		}
		return fmt.Sprintf("%v", v), nil
	case time.Time:
		if format != "" {
			return v.Format(format), nil
		}
		return v, nil
	case *time.Time:
		if v != nil {
			if format != "" {
				return v.Format(format), nil
			}
			return *v, nil
		}
		return "", nil
	case string:
		// Skip formatting for string values, even if format is specified
		// This prevents format conflicts (e.g., "Total" being formatted as date)
		return v, nil
	case int, int8, int16, int32, int64:
		return v, nil
	case uint, uint8, uint16, uint32, uint64:
		return v, nil
	case float32, float64:
		return v, nil
	case bool:
		return v, nil
	default:
		if format != "" {
			var err error
			value, err = formatValue(value, format)
			if err != nil {
				return "", err
			}
		}
		return fmt.Sprintf("%v", value), nil
	}
}
