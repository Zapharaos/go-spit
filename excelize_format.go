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
)
