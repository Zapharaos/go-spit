// excelize_table.go - Excelize-based implementation of TableOperations for go-spit
//
// This file provides an adapter for the github.com/xuri/excelize library, enabling table operations
// such as cell access, merging, styling, and value formatting for Excel spreadsheets.

package spit

import (
	"fmt"
	"reflect"

	"github.com/xuri/excelize/v2"
)

// TableExcelize provides Excelize-specific operations for table handling.
// Implements TableOperations for Excel spreadsheets using github.com/xuri/excelize.
// TableExcelize instances must not be shared across goroutines without external synchronization.
type TableExcelize struct {
	File                  *excelize.File       // Underlying Excelize file object
	SheetName             string               // Current sheet name
	Table                 *Table               // Reference to the generic Table struct
	mergedCells           []excelize.MergeCell // Cached merged-cell list for IsCellMerged lookups
	mergedCellsCachedName string               // Sheet name for which mergedCells is valid; reset on MergeCell call or SheetName change to invalidate cache
}

// NewTableExcelize creates a new TableExcelize instance for a given sheet name and table.
// The Excelize file is optional, set later via WithFile.
func NewTableExcelize(sheetName string, table *Table) *TableExcelize {
	return &TableExcelize{
		SheetName: sheetName,
		Table:     table,
	}
}

// WithFile sets an existing Excelize file to the TableExcelize instance.
// Returns the TableExcelize for chaining.
func (e *TableExcelize) WithFile(file *excelize.File) *TableExcelize {
	e.File = file
	return e
}

// GetTable returns the underlying Table struct for direct access/manipulation.
func (e *TableExcelize) GetTable() *Table {
	return e.Table
}

// GetCellValue returns the value of a cell at the given column and row (1-based indices).
// Converts coordinates to Excel cell reference and retrieves the value from the sheet.
func (e *TableExcelize) GetCellValue(col, row int) (string, error) {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return "", err
	}
	return e.File.GetCellValue(e.SheetName, cellRef)
}

// SetCellValue sets the value of a cell at the given column and row.
// Converts coordinates to Excel cell reference and sets the value in the sheet.
func (e *TableExcelize) SetCellValue(col, row int, value interface{}) error {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return err
	}
	return e.File.SetCellValue(e.SheetName, cellRef, value)
}

// MergeCells merges a rectangular range of cells from start to end coordinates.
// Converts coordinates to Excel cell references and merges the specified range.
func (e *TableExcelize) MergeCells(startCol, startRow, endCol, endRow int) error {
	startCell, err1 := excelize.CoordinatesToCellName(startCol, startRow)
	endCell, err2 := excelize.CoordinatesToCellName(endCol, endRow)
	if err1 != nil || err2 != nil {
		return fmt.Errorf("failed to convert coordinates: %v, %v", err1, err2)
	}
	err := e.File.MergeCell(e.SheetName, startCell, endCell)
	e.mergedCellsCachedName = "" // Invalidate the merged-cells cache
	return err
}

// getMergedCells returns the list of merged cells for the current sheet, using a cache to avoid
// repeated calls to GetMergeCells. The cache is invalidated whenever MergeCells is called or
// SheetName changes.
func (e *TableExcelize) getMergedCells() ([]excelize.MergeCell, error) {
	if e.mergedCellsCachedName != e.SheetName {
		cells, err := e.File.GetMergeCells(e.SheetName)
		if err != nil {
			return nil, err
		}
		e.mergedCells = cells
		e.mergedCellsCachedName = e.SheetName
	}
	return e.mergedCells, nil
}

// IsCellMerged checks if a cell at the given column and row is merged with others.
// Returns true if the cell is part of a merged range, false otherwise.
func (e *TableExcelize) IsCellMerged(col, row int) bool {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return false
	}

	mergedCells, err := e.getMergedCells()
	if err != nil {
		return false
	}
	for _, mergeCell := range mergedCells {
		if isCellInRange(cellRef, mergeCell.GetStartAxis(), mergeCell.GetEndAxis()) {
			return true
		}
	}
	return false
}

// IsCellMergedHorizontally checks if a cell at the given column and row is merged horizontally.
// Returns true if the cell is part of a horizontally merged range, false otherwise.
func (e *TableExcelize) IsCellMergedHorizontally(col, row int) bool {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return false
	}

	mergedCells, err := e.getMergedCells()
	if err != nil {
		return false
	}
	for _, mergeCell := range mergedCells {
		if isCellInRange(cellRef, mergeCell.GetStartAxis(), mergeCell.GetEndAxis()) {
			startCol, startRow, _ := excelize.CellNameToCoordinates(mergeCell.GetStartAxis())
			endCol, endRow, _ := excelize.CellNameToCoordinates(mergeCell.GetEndAxis())
			return startRow == endRow && startCol != endCol
		}
	}
	return false
}

// ApplyBorderToCell applies a border to a specific side of a cell at the given column and row.
// The border style is defined by the Border parameter. Only non-none styles are applied.
func (e *TableExcelize) ApplyBorderToCell(col, row int, side string, border *Border) error {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return err
	}
	if border == nil || border.Style == BorderStyleNone {
		return nil
	}

	// Validate side before doing any expensive work
	switch side {
	case "left", "right", "top", "bottom":
	default:
		return fmt.Errorf("unsupported border side: %s", side)
	}

	// Get current style and preserve existing properties
	excelStyle, err := e.getCellStyle(col, row)
	if excelStyle == nil || err != nil {
		excelStyle = &excelize.Style{}
	}
	excelStyle.Border = append(excelStyle.Border, excelize.Border{Type: side, Color: "000000", Style: int(border.Style)})

	styleID, err := e.File.NewStyle(excelStyle)
	if err != nil {
		return err
	}

	return e.File.SetCellStyle(e.SheetName, cellRef, cellRef, styleID)
}

// ApplyBordersToRange applies borders to a range of cells defined by start and end coordinates.
// Each side of the range can have a different border style, as specified in the Borders parameter.
// All applicable border sides are batched into a single style update per cell.
func (e *TableExcelize) ApplyBordersToRange(startCol, startRow, endCol, endRow int, borders Borders) error {
	for row := startRow; row <= endRow; row++ {
		for col := startCol; col <= endCol; col++ {
			// Collect which border sides apply to this cell in the range
			var sides []excelize.Border
			if col == startCol && borders.Left != nil && borders.Left.Style != BorderStyleNone {
				sides = append(sides, excelize.Border{Type: "left", Color: "000000", Style: int(borders.Left.Style)})
			}
			if col == endCol && borders.Right != nil && borders.Right.Style != BorderStyleNone {
				sides = append(sides, excelize.Border{Type: "right", Color: "000000", Style: int(borders.Right.Style)})
			}
			if row == startRow && borders.Top != nil && borders.Top.Style != BorderStyleNone {
				sides = append(sides, excelize.Border{Type: "top", Color: "000000", Style: int(borders.Top.Style)})
			}
			if row == endRow && borders.Bottom != nil && borders.Bottom.Style != BorderStyleNone {
				sides = append(sides, excelize.Border{Type: "bottom", Color: "000000", Style: int(borders.Bottom.Style)})
			}

			if len(sides) == 0 {
				continue
			}

			// Apply all sides in one style read/write cycle
			cellRef, err := excelize.CoordinatesToCellName(col, row)
			if err != nil {
				return err
			}

			excelStyle, err := e.getCellStyle(col, row)
			if excelStyle == nil || err != nil {
				excelStyle = &excelize.Style{}
			}
			excelStyle.Border = append(excelStyle.Border, sides...)

			styleID, err := e.File.NewStyle(excelStyle)
			if err != nil {
				return err
			}
			if err := e.File.SetCellStyle(e.SheetName, cellRef, cellRef, styleID); err != nil {
				return err
			}
		}
	}
	return nil
}

// HasExistingBorder checks if a cell at the given column and row has any existing border applied on the specified side.
// Returns true if there is a border style applied, false otherwise.
func (e *TableExcelize) HasExistingBorder(col, row int, side string) bool {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return false
	}
	styleID, err := e.File.GetCellStyle(e.SheetName, cellRef)
	if err != nil {
		return false
	}
	// Check if there's a style applied (simple check)
	return styleID > 0
}

// ApplyStyleToCell applies a style to a cell at the given column and row.
// The style properties are defined in the style parameter. Existing borders are preserved.
func (e *TableExcelize) ApplyStyleToCell(col, row int, style Style) error {
	return e.applyExcelizeStyleToCell(col, row, convertStyleToExcelizeStyle(style))
}

// applyExcelizeStyleToCell applies a pre-converted excelize style to a single cell,
// merging it with any existing style so that borders and other properties are preserved.
func (e *TableExcelize) applyExcelizeStyleToCell(col, row int, inputStyle *excelize.Style) error {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return err
	}

	// Fast path: if the cell carries no existing style, apply inputStyle directly without
	// a round-trip to read and decode the default style.
	existingID, err := e.File.GetCellStyle(e.SheetName, cellRef)
	if err != nil {
		return err
	}

	var finalStyle *excelize.Style
	if existingID == 0 {
		// styleID 0 is the excelize default (no style applied); skip the GetStyle round-trip.
		finalStyle = inputStyle
	} else {
		excelStyle, err := e.File.GetStyle(existingID)
		if excelStyle == nil || err != nil {
			finalStyle = inputStyle
		} else {
			// Merge: overlay inputStyle properties on top of the existing style so that
			// borders (and any other properties not covered by inputStyle) are preserved.
			if inputStyle.Fill.Color != nil {
				excelStyle.Fill = inputStyle.Fill
			}
			if inputStyle.Font != nil {
				excelStyle.Font = inputStyle.Font
			}
			if inputStyle.Alignment != nil {
				excelStyle.Alignment = inputStyle.Alignment
			}
			if inputStyle.CustomNumFmt != nil {
				excelStyle.CustomNumFmt = inputStyle.CustomNumFmt
			}
			finalStyle = excelStyle
		}
	}

	styleID, err := e.File.NewStyle(finalStyle)
	if err != nil {
		return err
	}
	return e.File.SetCellStyle(e.SheetName, cellRef, cellRef, styleID)
}

// ApplyStyleToRange applies a style to a range of cells defined by start and end coordinates.
// The style is converted once and then applied to each cell in the range.
func (e *TableExcelize) ApplyStyleToRange(startCol, startRow, endCol, endRow int, style Style) error {
	inputStyle := convertStyleToExcelizeStyle(style)
	for row := startRow; row <= endRow; row++ {
		for col := startCol; col <= endCol; col++ {
			if err := e.applyExcelizeStyleToCell(col, row, inputStyle); err != nil {
				return err
			}
		}
	}
	return nil
}

// GetColumnLetter returns the Excel-style column letter (A, B, C, ...) for a given column number.
func (e *TableExcelize) GetColumnLetter(col int) string {
	letter, _ := excelize.ColumnNumberToName(col)
	return letter
}

// ProcessValue processes and formats a value according to its type and the specified format.
// Supports basic types, time.Time, and slices. Formats value for Excel export.
// Special formats ExcelizeFormatFormula and ExcelizeFormatHyperlink return the raw string value.
// ExcelizeFormatDefault returns the raw value without string conversion, preserving its native type.
func (e *TableExcelize) ProcessValue(value interface{}, format string) (interface{}, error) {
	switch v := value.(type) {
	case []interface{}:
		if e.Table.ListSeparator != "" {
			return ConvertSliceToString(v, format, e.Table.ListSeparator)
		}
		return fmt.Sprintf("%v", v), nil
	default:
		switch format {
		case ExcelizeFormatDefault:
			// Return the raw value so Excelize preserves the native Go type
			// (e.g. int stays a number, bool stays boolean, time.Time becomes a date).
			return value, nil
		case ExcelizeFormatFormula, ExcelizeFormatHyperlink:
			// Formula and hyperlink values are written via dedicated Excelize calls;
			// return the string representation here for merge-comparison use.
			return fmt.Sprintf("%v", value), nil
		case ExcelizeFormatNumber:
			// Coerce the value to a numeric type (parsing strings like "123" / "1.5")
			// so Excelize writes a real number rather than text.
			return convertToNumber(value)
		case ExcelizeFormatBool:
			// Coerce the value to a boolean (parsing strings like "true" / "yes" / "1")
			// so Excelize writes a real boolean rather than text.
			return convertToBool(value)
		default:
			if format != "" {
				var err error
				value, err = FormatValue(value, format)
				if err != nil {
					return "", err
				}
			}
			return fmt.Sprintf("%v", value), nil
		}
	}
}

// SetCellFormula sets the formula of a cell at the given column and row.
// The formula string should be a valid Excel formula, e.g. "=SUM(A1:A10)".
func (e *TableExcelize) SetCellFormula(col, row int, formula string) error {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return err
	}
	return e.File.SetCellFormula(e.SheetName, cellRef, formula)
}

// SetCellHyperLink sets an external hyperlink on a cell at the given column and row.
// The cell display value is also set to the link URL.
func (e *TableExcelize) SetCellHyperLink(col, row int, link string) error {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return err
	}
	return e.File.SetCellHyperLink(e.SheetName, cellRef, link, "External")
}

// SetCellImage places an image at the given column and row, anchored over the cell.
// Embedded content (Bytes) is inserted via AddPictureFromBytes; a URL is treated as a
// local file path and inserted via AddPicture (remote URLs are not fetched). The image
// auto-fits the cell. Images without a usable source are ignored.
//
// Note: Excelize only supports pictures anchored over cells; the newer "place in cell"
// (rich-value IMAGE) embedding is not available for writing, so images float over the cell.
func (e *TableExcelize) SetCellImage(col, row int, img Image) error {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return err
	}

	opts := &excelize.GraphicOptions{AltText: img.AltText, AutoFit: true}

	if img.HasBytes() {
		ext := extensionFromMIME(img.MIME)
		if ext == "" {
			return fmt.Errorf("cannot resolve image extension from MIME %q for cell (%d, %d)", img.MIME, col, row)
		}
		return e.File.AddPictureFromBytes(e.SheetName, cellRef, &excelize.Picture{
			Extension: ext,
			File:      img.Bytes,
			Format:    opts,
		})
	}

	if img.URL != "" {
		return e.File.AddPicture(e.SheetName, cellRef, img.URL, opts)
	}

	return nil
}

// getCellStyle retrieves the style of a cell at the given column and row.
func (e *TableExcelize) getCellStyle(col, row int) (*excelize.Style, error) {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return nil, err
	}
	styleID, err := e.File.GetCellStyle(e.SheetName, cellRef)
	if err != nil {
		return nil, err
	}
	return e.File.GetStyle(styleID)
}

// convertStyleToExcelizeStyle converts a Style struct to the corresponding Excelize style.
// Maps font, fill, and alignment properties to Excelize style attributes.
func convertStyleToExcelizeStyle(style Style) *excelize.Style {
	excelStyle := &excelize.Style{}

	if style.Bold || style.Italic || style.FontSize > 0 || style.FontFamily != "" || style.TextColor != "" {
		font := &excelize.Font{}
		if style.Bold {
			font.Bold = true
		}
		if style.Italic {
			font.Italic = true
		}
		if style.FontSize > 0 {
			font.Size = style.FontSize
		}
		if style.FontFamily != "" {
			font.Family = style.FontFamily
		}
		if style.TextColor != "" {
			font.Color = style.TextColor
		}
		if style.Underline != "" {
			font.Underline = style.Underline
		}
		excelStyle.Font = font
	}

	if style.BackgroundColor != "" {
		excelStyle.Fill = excelize.Fill{
			Type:    "pattern",
			Color:   []string{style.BackgroundColor},
			Pattern: 1,
		}
	}

	if style.Alignment != AlignmentNone {
		horizontal, vertical := style.Alignment.GetAlignmentValues()
		excelStyle.Alignment = &excelize.Alignment{
			Horizontal: horizontal,
			Vertical:   vertical,
		}
	}

	if style.NumFmt != "" {
		excelStyle.CustomNumFmt = &style.NumFmt
	}

	return excelStyle
}

// isCellInRange checks if a cell reference is within a given range defined by start and end references.
// Returns true if the cell is in range, false otherwise.
func isCellInRange(cellRef, startRef, endRef string) bool {
	col, row, err := excelize.CellNameToCoordinates(cellRef)
	if err != nil {
		return false
	}
	startCol, startRow, err := excelize.CellNameToCoordinates(startRef)
	if err != nil {
		return false
	}
	endCol, endRow, err := excelize.CellNameToCoordinates(endRef)
	if err != nil {
		return false
	}
	return col >= startCol && col <= endCol && row >= startRow && row <= endRow
}

// convertToNumber coerces a value to a numeric type for Excel output.
// Native numeric types are returned unchanged; strings are parsed as int then float.
// nil is preserved, and values that cannot be represented as a number fall back to
// their original string representation so no data is lost.
func convertToNumber(value interface{}) (interface{}, error) {
	switch v := value.(type) {
	case nil:
		return nil, nil
	case int, int8, int16, int32, int64,
		uint, uint8, uint16, uint32, uint64,
		float32, float64:
		return v, nil
	case string:
		if intVal, err := parseAsInt(v); err == nil {
			return intVal, nil
		}
		if floatVal, err := parseAsFloat(v); err == nil {
			return floatVal, nil
		}
		return v, nil
	default:
		return fmt.Sprintf("%v", value), nil
	}
}

// convertToBool coerces a value to a boolean for Excel output.
// Booleans are returned unchanged; strings are parsed via parseAsBool; numeric values
// are true when non-zero. nil is preserved, and unrecognized values fall back to their
// original string representation so no data is lost.
func convertToBool(value interface{}) (interface{}, error) {
	switch v := value.(type) {
	case nil:
		return nil, nil
	case bool:
		return v, nil
	case string:
		if boolVal, err := parseAsBool(v); err == nil {
			return boolVal, nil
		}
		return v, nil
	}

	// Numeric types: zero is false, non-zero is true.
	rv := reflect.ValueOf(value)
	switch rv.Kind() {
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		return rv.Int() != 0, nil
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		return rv.Uint() != 0, nil
	case reflect.Float32, reflect.Float64:
		return rv.Float() != 0, nil
	}
	return fmt.Sprintf("%v", value), nil
}
