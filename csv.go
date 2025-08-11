// csv.go - CSV export logic.
//
// This file provides functions to write tabular data to CSV files.

package go_spit

import (
	stdcsv "encoding/csv"
	"fmt"
	"io"
)

// CSV contains CSV-specific export parameters
type CSV struct {
	writer    *stdcsv.Writer // Private CSV writer instance
	Separator string         // Separator used for CSV fields, default is comma
	table     *Table         // Reference to the Table being exported
}

// NewCsv creates a new CSV instance with the specified separator and table
func NewCsv(separator string, t *Table) *CSV {
	return &CSV{
		Separator: separator,
		table:     t,
	}
}

// WriteDataToFile writes generic data to file using the generic file writer
func (csv *CSV) WriteDataToFile(options FileWriteOptions) (*FileWriteResult, error) {
	L().Info("Starting CSV export to file", String("filename", options.Filename))

	// Ensure extension is set for CSV files
	if options.extension == "" {
		options.extension = FormatCSV.String()
	}

	// Create a write function that handles the CSV file creation and writing
	writeFunc := func(writer io.Writer) error {
		csv.writer = stdcsv.NewWriter(writer)
		return csv.writeData()
	}

	// Use the generic file writer to handle the actual file writing
	result, err := options.writeToFile(writeFunc)
	if err != nil {
		L().Error("Failed to write CSV to file", Error(err))
		return nil, err
	}

	L().Info("CSV export completed", String("filename", options.Filename))
	return result, nil
}

// writeData writes the provided data to the CSV writer
func (csv *CSV) writeData() error {
	L().Debug("Writing data to CSV...")
	// Set the CSV delimiter (comma by default)
	if csv.writer.Comma == 0 {
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
			// lookup the value for this column in the current row
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
func (csv *CSV) writeHeaders() error {
	maxDepth := csv.table.Columns.getMaxDepth()
	totalCols := csv.table.Columns.getTotalColumnCount()
	L().Debug("Writing header levels", Int("levels", maxDepth), Int("columns", totalCols))

	// Generate header rows for each level
	for level := 0; level < maxDepth; level++ {
		headerRow := make([]string, totalCols)
		csv.fillHeaderLevel(headerRow, level, 0, 0)
		if err := csv.writer.Write(headerRow); err != nil {
			return fmt.Errorf("error writing header row: %w", err)
		}
	}
	L().Debug("CSV headers written successfully.")
	return nil
}

// fillHeaderLevel recursively fills a header row for a specific level
func (csv *CSV) fillHeaderLevel(headerRow []string, targetLevel int, currentLevel int, colIndex int) int {
	for _, column := range csv.table.Columns {
		if currentLevel == targetLevel {
			// This is the level we want to fill
			if column.hasSubColumns() {
				// For parent columns, write the label and span across all sub-columns
				colSpan := column.getColumnCount()
				headerRow[colIndex] = column.Label
				// Fill the span with empty strings for merged appearance
				for i := 1; i < colSpan; i++ {
					if colIndex+i < len(headerRow) {
						headerRow[colIndex+i] = ""
					}
				}
				colIndex += colSpan
			} else {
				// Leaf column at this level
				headerRow[colIndex] = column.Label
				colIndex++
			}
		} else if currentLevel < targetLevel {
			// We need to go deeper
			if column.hasSubColumns() {
				colIndex = csv.fillHeaderLevel(headerRow, targetLevel, currentLevel+1, colIndex)
			} else {
				// Leaf column, but we're looking for a deeper level - leave empty
				if targetLevel > currentLevel {
					headerRow[colIndex] = ""
				}
				colIndex++
			}
		}
	}
	return colIndex
}

// processValue processes a value based on its type and format
func (csv *CSV) processValue(value interface{}, format string) (string, error) {
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
