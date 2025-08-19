// csv.go - CSV export logic.
//
// This file provides functions to write tabular data to CSV files.

package spit

import (
	stdcsv "encoding/csv"
	"fmt"
	"io"
)

// ExportCSV writes generic table data to a CSV file using the generic file writer.
func ExportCSV(separator string, t *Table, params FileWriteParams) (*FileWriteResult, error) {
	// Ensure extension is set for CSV files
	if params.extension == "" {
		params.extension = FormatCSV.String()
	}

	csvConfig := &csv{
		separator: separator,
		table:     t,
		params:    params,
	}

	L().Info("Starting CSV export to file", String("filename", csvConfig.params.Filename))

	// Create a write function that handles the CSV file creation and writing
	writeFunc := func(writer io.Writer) error {
		csvConfig.writer = stdcsv.NewWriter(writer)
		return csvConfig.writeData()
	}

	// Use the generic file writer to handle the actual file writing
	result, err := csvConfig.params.writeToFile(writeFunc)
	if err != nil {
		L().Error("Failed to write CSV to file", Error(err))
		return nil, err
	}

	L().Info("CSV export completed", String("filename", csvConfig.params.Filename))
	return result, nil
}

// csv contains CSV-specific export parameters and logic.
type csv struct {
	writer    *stdcsv.Writer  // Private CSV writer instance
	separator string          // Separator used for CSV fields, default is comma
	table     *Table          // Reference to the Table being exported
	params    FileWriteParams // File write parameters for the CSV export
}

// writeData writes the provided table data to the CSV writer.
// Handles headers, data rows, and value formatting.
func (csv *csv) writeData() error {
	L().Debug("Writing data to CSV...")

	// Set the CSV delimiter (comma by default)
	if csv.separator != "" {
		csv.writer.Comma = rune(csv.separator[0])
	} else {
		csv.writer.Comma = ','
	}

	// Write headers if requested
	if csv.table.WriteHeader && len(csv.table.Columns) > 0 {
		L().Debug("Writing CSV headers...")
		if err := csv.writeHeaders(); err != nil {
			return fmt.Errorf("error writing CSV headers: %w", err)
		}
	}

	// Get flattened columns for data processing
	flatColumns := csv.table.Columns.getFlattenedColumns()

	// Write each data row to the CSV
	for rowIdx, item := range csv.table.Data {
		record := make([]string, 0, len(flatColumns))
		for _, column := range flatColumns {
			// Lookup the value for this column in the current row
			value, err := item.lookup(column.Name)
			if err != nil {
				return fmt.Errorf("missing value for column %s in row %d: %w", column.Name, rowIdx, err)
			}

			// Process the value based on column format (e.g., date, number)
			processedValue, err := csv.processValue(value, column.Format)
			if err != nil {
				return fmt.Errorf("error processing value for column %s in row %d: %w", column.Name, rowIdx, err)
			}
			record = append(record, processedValue)
		}

		// Write the processed record to the CSV file
		if err := csv.writer.Write(record); err != nil {
			return fmt.Errorf("error writing CSV record for row %d: %w", rowIdx, err)
		}
	}

	// Flush buffered data to the underlying writer
	csv.writer.Flush()
	if err := csv.writer.Error(); err != nil {
		return fmt.Errorf("error flushing CSV writer: %w", err)
	}

	L().Debug("CSV data writing complete.")
	return nil
}

// writeHeaders writes header rows to represent the hierarchical column structure
// Each row corresponds to a level in the column hierarchy, allowing for grouped headers in the CSV output.
func (csv *csv) writeHeaders() error {
	maxDepth := csv.table.Columns.getMaxDepth()
	totalCols := csv.table.Columns.getTotalColumnCount()
	L().Debug("Writing header levels", Int("levels", maxDepth), Int("columns", totalCols))

	// Generate header rows for each level
	for level := 0; level < maxDepth; level++ {
		headerRow := make([]string, totalCols)
		csv.fillHeaderLevel(headerRow, level, 0, 0, csv.table.Columns)
		if err := csv.writer.Write(headerRow); err != nil {
			return fmt.Errorf("error writing header row: %w", err)
		}
	}
	L().Debug("CSV headers written successfully.")
	return nil
}

// fillHeaderLevel recursively fills a header row for a specific level using the provided columns.
// Handles parent columns (spanning multiple sub-columns) and leaf columns.
func (csv *csv) fillHeaderLevel(headerRow []string, targetLevel int, currentLevel int, colIndex int, columns Columns) int {
	for _, column := range columns {
		if currentLevel == targetLevel {
			// This is the level we want to fill
			if column.hasSubColumns() {
				// For parent columns, write the label and span across all sub-columns
				colSpan := column.getColumnCount()
				if colIndex < len(headerRow) {
					headerRow[colIndex] = column.Label
				}
				// Fill the span with empty strings for merged appearance
				for i := 1; i < colSpan; i++ {
					if colIndex+i < len(headerRow) {
						headerRow[colIndex+i] = ""
					}
				}
				colIndex += colSpan
			} else {
				// Leaf column at this level
				if colIndex < len(headerRow) {
					headerRow[colIndex] = column.Label
				}
				colIndex++
			}
		} else if currentLevel < targetLevel {
			// We need to go deeper
			if column.hasSubColumns() {
				// Recurse into sub-columns
				colIndex = csv.fillHeaderLevel(headerRow, targetLevel, currentLevel+1, colIndex, column.Columns)
			} else {
				// Leaf column, but we're looking for a deeper level - leave empty
				if colIndex < len(headerRow) {
					headerRow[colIndex] = ""
				}
				colIndex++
			}
		}
	}
	return colIndex
}

// processValue processes a value based on its type and format for CSV output.
// Handles slices, formatting, and string conversion.
func (csv *csv) processValue(value interface{}, format string) (string, error) {
	switch v := value.(type) {
	case []interface{}:
		if csv.table.ListSeparator != "" {
			return convertSliceToString(v, format, csv.table.ListSeparator)
		}
	default:
		if format != "" {
			var err error
			value, err = formatValue(value, format)
			if err != nil {
				return "", err
			}
		}
	}
	// Convert value to string
	return fmt.Sprintf("%v", value), nil
}
