// excelize_format.go - Excelize-specific cell content format constants.
//
// This file defines format constants for special Excelize cell content types.
// Use these as the Format field on a Column to control how cell values are written to Excel.

package spit

const (
	// ExcelizeFormatDefault passes the raw value to Excelize without string conversion,
	// allowing Excelize to use the native Go type (number, bool, time, etc.).
	ExcelizeFormatDefault = "default"

	// ExcelizeFormatFormula sets the cell content as an Excel formula using SetCellFormula.
	// The column value must be a formula string, e.g. "=SUM(A1:A10)".
	ExcelizeFormatFormula = "formula"

	// ExcelizeFormatHyperlink sets the cell as a clickable external hyperlink using SetCellHyperLink.
	// The column value must be a URL string, e.g. "https://example.com".
	// The cell display text is also set to the URL value.
	ExcelizeFormatHyperlink = "hyperlink"

	// ExcelizeFormatNumber coerces the cell value to a numeric type before writing.
	// String values such as "123" or "1.5" are parsed to int/float so Excelize stores
	// a real number; unparseable values fall back to their string representation.
	ExcelizeFormatNumber = "number"

	// ExcelizeFormatBool coerces the cell value to a boolean before writing.
	// String values such as "true"/"false", "yes"/"no", or "1"/"0" are parsed so
	// Excelize stores a real boolean; unparseable values fall back to their string representation.
	ExcelizeFormatBool = "bool"
)
