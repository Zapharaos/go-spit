package go_spit

import (
	stdcsv "encoding/csv"
	"fmt"
	"io"
	"unicode/utf8"

	"github.com/Zapharaos/go-spit/internal/file"
	"github.com/Zapharaos/go-spit/internal/table"
	"github.com/Zapharaos/go-spit/internal/utils"
)

// CSV contains CSV-specific export parameters
type CSV struct {
	Writer    *stdcsv.Writer
	Separator string
	table     *table.Table
}

// NewCsv creates a new CSV instance with the specified separator and table
func NewCsv(separator string, t *table.Table) *CSV {
	return &CSV{
		Separator: separator,
		table:     t,
	}
}

// WriteDataToFile writes generic data to file using the generic file writer
func (csv *CSV) WriteDataToFile(options file.WriteOptions) (*file.WriteResult, error) {
	// Ensure extension is set for CSV files
	if options.Extension == "" {
		options.Extension = FormatCSV.String()
	}

	writeFunc := func(writer io.Writer) error {
		csv.Writer = stdcsv.NewWriter(writer)
		return csv.writeData()
	}

	return options.WriteToFile(writeFunc)
}

// writeData writes the provided data to the CSV writer
func (csv *CSV) writeData() error {
	if len(csv.Separator) == 1 {
		csv.Writer.Comma, _ = utf8.DecodeRune([]byte(csv.Separator))
		if csv.Writer.Comma == utf8.RuneError {
			csv.Writer.Comma = ','
		}
	} else {
		csv.Writer.Comma = ','
	}

	// Write multi-level headers if requested
	if csv.table.WriteHeader && len(csv.table.Columns) > 0 {
		if err := csv.writeMultiLevelHeaders(); err != nil {
			return err
		}
	}

	// Get flattened columns for data processing
	flatColumns := csv.table.Columns.GetFlattenedColumns()

	// Write data rows
	for _, item := range csv.table.Data {
		record := make([]string, 0, len(flatColumns))
		for _, column := range flatColumns {
			value, err := item.Lookup(column.Name)
			if err != nil {
				record = append(record, "")
				continue
			}

			// Process the value based on column format
			processedValue, err := csv.processValue(value, column.Format)
			if err != nil {
				return fmt.Errorf("error processing value for column %s: %w", column.Name, err)

			}
			record = append(record, processedValue)
		}

		if err := csv.Writer.Write(record); err != nil {
			return err
		}
	}

	csv.Writer.Flush()
	return csv.Writer.Error()
}

// writeMultiLevelHeaders writes multiple header rows to represent the hierarchical column structure
func (csv *CSV) writeMultiLevelHeaders() error {
	maxDepth := csv.table.Columns.GetMaxDepth()
	totalCols := csv.table.Columns.GetTotalColumnCount()

	// Generate header rows for each level
	for level := 0; level < maxDepth; level++ {
		headerRow := make([]string, totalCols)
		csv.fillHeaderLevel(headerRow, level, 0, 0)

		if err := csv.Writer.Write(headerRow); err != nil {
			return err
		}
	}

	return nil
}

// fillHeaderLevel recursively fills a header row for a specific level
func (csv *CSV) fillHeaderLevel(headerRow []string, targetLevel int, currentLevel int, colIndex int) int {
	for _, column := range csv.table.Columns {
		if currentLevel == targetLevel {
			// This is the level we want to fill
			if column.HasSubColumns() {
				// For parent columns, write the label and span across all sub-columns
				colSpan := column.GetColumnCount()
				headerRow[colIndex] = column.Label
				// Fill the span with empty strings or the same label (depending on preference)
				for i := 1; i < colSpan; i++ {
					if colIndex+i < len(headerRow) {
						headerRow[colIndex+i] = "" // Empty for merged appearance
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
			if column.HasSubColumns() {
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
			return utils.ConvertSliceToString(v, format, csv.table.ListSeparator)
		}
	default:
		if format != "" {
			var err error
			value, err = utils.FormatValue(value, format)
			if err != nil {
				return "", err
			}
		}
	}

	return fmt.Sprintf("%v", value), nil
}
