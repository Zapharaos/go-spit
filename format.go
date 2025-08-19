// format.go - Export formats.
//
// This file defines the Format type and related constants for supported export formats.
// Used to select and identify the output format for exported files (e.g., CSV, XLSX).

package spit

import "fmt"

// Format represents the export format for files (e.g., CSV, XLSX).
type Format uint8

const (
	FormatUnknown Format = iota // Unknown format (default)
	FormatCSV                   // CSV format
	FormatXSLX                  // XLSX format
)

// formats maps Format values to their string representations.
var formats = map[Format]string{
	FormatCSV:  "csv",
	FormatXSLX: "xlsx",
}

// String returns the string representation of the Format.
// If the format is not recognized, returns a generic string with the format value.
func (f Format) String() string {
	if str, ok := formats[f]; ok {
		return str
	}
	return fmt.Sprintf("Format(%d)", f)
}
