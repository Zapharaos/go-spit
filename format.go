package go_spit

import "fmt"

type Format uint8

const (
	FormatUnknown Format = iota
	FormatCSV
	FormatXSLX
)

var formats = map[Format]string{
	FormatCSV:  "csv",
	FormatXSLX: "xlsx",
}

// String returns the string representation of the Format
func (f Format) String() string {
	if str, ok := formats[f]; ok {
		return str
	}
	return fmt.Sprintf("Format(%d)", f)
}
