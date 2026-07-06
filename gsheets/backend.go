package gsheets

import (
	"fmt"
	"strconv"

	spit "github.com/Zapharaos/go-spit"
	"google.golang.org/api/sheets/v4"
)

// gsheetTable implements spit.TableOperations on top of an in-memory grid of Google
// Sheets CellData, then serializes the grid into batchUpdate requests.
type gsheetTable struct {
	table   *spit.Table
	sheetID int64
	cells   map[int]map[int]*sheets.CellData // cells[row][col], both 1-based
	merges  []*sheets.GridRange
	maxRow  int
	maxCol  int
}

// Ensure the backend satisfies the shared interface.
var _ spit.TableOperations = (*gsheetTable)(nil)

func newGSheetTable(t *spit.Table, sheetID int64) *gsheetTable {
	return &gsheetTable{
		table:   t,
		sheetID: sheetID,
		cells:   make(map[int]map[int]*sheets.CellData),
	}
}

// ---- Grid helpers -----------------------------------------------------------

func (g *gsheetTable) cell(col, row int) *sheets.CellData {
	if g.cells[row] == nil {
		g.cells[row] = make(map[int]*sheets.CellData)
	}
	c := g.cells[row][col]
	if c == nil {
		c = &sheets.CellData{}
		g.cells[row][col] = c
	}
	if row > g.maxRow {
		g.maxRow = row
	}
	if col > g.maxCol {
		g.maxCol = col
	}
	return c
}

func (g *gsheetTable) peek(col, row int) *sheets.CellData {
	if g.cells[row] == nil {
		return nil
	}
	return g.cells[row][col]
}

// format returns the cell's UserEnteredFormat, allocating it on first use so that
// styles and borders accumulate on the same cell.
func (g *gsheetTable) format(col, row int) *sheets.CellFormat {
	c := g.cell(col, row)
	if c.UserEnteredFormat == nil {
		c.UserEnteredFormat = &sheets.CellFormat{}
	}
	return c.UserEnteredFormat
}

// gridRange converts 1-based inclusive coordinates to a 0-based, end-exclusive GridRange.
func (g *gsheetTable) gridRange(startCol, startRow, endCol, endRow int) *sheets.GridRange {
	return &sheets.GridRange{
		SheetId:          g.sheetID,
		StartRowIndex:    int64(startRow - 1),
		EndRowIndex:      int64(endRow),
		StartColumnIndex: int64(startCol - 1),
		EndColumnIndex:   int64(endCol),
		ForceSendFields:  []string{"StartRowIndex", "StartColumnIndex"},
	}
}

// ---- Build pipeline ---------------------------------------------------------

// build populates the grid (preamble, headers, data) and applies the shared merging
// and styling pipelines, mirroring the XLSX/HTML write flow.
func (g *gsheetTable) build() error {
	t := g.table
	currentRow := 1

	if len(t.Preamble) > 0 {
		for i, row := range t.Preamble {
			for j, val := range row.Values {
				if err := g.SetCellValue(j+1, currentRow+i, val); err != nil {
					return err
				}
			}
		}
		currentRow += len(t.Preamble)
	}

	if t.WriteHeader && len(t.Columns) > 0 {
		n, err := g.writeHeaders(currentRow)
		if err != nil {
			return err
		}
		currentRow += n
	}

	flat := t.Columns.GetFlattenedColumns()
	for _, item := range t.Data {
		col := 1
		for _, column := range flat {
			if err := g.writeCell(item, column, col, currentRow); err != nil {
				return err
			}
			col++
		}
		currentRow++
	}

	if err := t.ProcessMerging(g); err != nil {
		return fmt.Errorf("process merging: %w", err)
	}
	if err := t.RenderStyles(g); err != nil {
		return fmt.Errorf("render styles: %w", err)
	}
	return nil
}

func (g *gsheetTable) writeHeaders(startRow int) (int, error) {
	t := g.table
	maxDepth := t.Columns.GetMaxDepth()
	if maxDepth == 1 {
		for i, column := range t.Columns {
			if err := g.SetCellValue(i+1, startRow, column.Label); err != nil {
				return 0, err
			}
		}
		return 1, nil
	}
	if err := g.writeHeaderRow(t.Columns, startRow, startRow+maxDepth-1, 1); err != nil {
		return 0, err
	}
	return maxDepth, nil
}

func (g *gsheetTable) writeHeaderRow(columns spit.Columns, currentRow, maxRow, startCol int) error {
	currentCol := startCol
	for _, column := range columns {
		if err := g.SetCellValue(currentCol, currentRow, column.Label); err != nil {
			return err
		}
		if column.HasSubColumns() {
			if currentRow < maxRow {
				if err := g.writeHeaderRow(column.Columns, currentRow+1, maxRow, currentCol); err != nil {
					return err
				}
			}
			currentCol += column.CountSubColumns()
		} else {
			currentCol++
		}
	}
	return nil
}

func (g *gsheetTable) writeCell(item spit.Data, column *spit.Column, col, row int) error {
	value, err, found := item.Lookup(column.Name)
	if err == nil && !found {
		return nil
	}
	if err != nil {
		return err
	}

	// Image values become =IMAGE() formulas (URL) or a text fallback.
	if img, ok := toImage(value); ok {
		return g.SetCellImage(col, row, img)
	}

	processed, err := g.ProcessValue(value, column.Format)
	if err != nil {
		return err
	}

	switch column.Format {
	case spit.ExcelizeFormatFormula:
		return g.SetCellFormula(col, row, fmt.Sprintf("%v", processed))
	case spit.ExcelizeFormatHyperlink:
		link := fmt.Sprintf("%v", processed)
		return g.SetCellHyperLink(col, row, link)
	default:
		// Keep unformatted numeric values native so Sheets treats them as numbers.
		if column.Format == "" && isNumeric(value) {
			return g.SetCellValue(col, row, value)
		}
		return g.SetCellValue(col, row, processed)
	}
}

// requests serializes the accumulated grid and merges into batchUpdate requests.
func (g *gsheetTable) requests() []*sheets.Request {
	var reqs []*sheets.Request

	if g.maxRow > 0 && g.maxCol > 0 {
		rows := make([]*sheets.RowData, 0, g.maxRow)
		for r := 1; r <= g.maxRow; r++ {
			values := make([]*sheets.CellData, 0, g.maxCol)
			for c := 1; c <= g.maxCol; c++ {
				if cell := g.peek(c, r); cell != nil {
					values = append(values, cell)
				} else {
					values = append(values, &sheets.CellData{})
				}
			}
			rows = append(rows, &sheets.RowData{Values: values})
		}
		reqs = append(reqs, &sheets.Request{UpdateCells: &sheets.UpdateCellsRequest{
			Rows:   rows,
			Fields: "userEnteredValue,userEnteredFormat",
			Start: &sheets.GridCoordinate{
				SheetId:         g.sheetID,
				RowIndex:        0,
				ColumnIndex:     0,
				ForceSendFields: []string{"RowIndex", "ColumnIndex"},
			},
		}})
	}

	for _, m := range g.merges {
		reqs = append(reqs, &sheets.Request{MergeCells: &sheets.MergeCellsRequest{
			Range:     m,
			MergeType: "MERGE_ALL",
		}})
	}
	return reqs
}

// ---- spit.TableOperations implementation ------------------------------------

func (g *gsheetTable) GetTable() *spit.Table { return g.table }

func (g *gsheetTable) SetCellValue(col, row int, value interface{}) error {
	g.cell(col, row).UserEnteredValue = extendedValue(value)
	return nil
}

func (g *gsheetTable) GetCellValue(col, row int) (string, error) {
	c := g.peek(col, row)
	if c == nil || c.UserEnteredValue == nil {
		return "", nil
	}
	ev := c.UserEnteredValue
	switch {
	case ev.StringValue != nil:
		return *ev.StringValue, nil
	case ev.NumberValue != nil:
		return strconv.FormatFloat(*ev.NumberValue, 'f', -1, 64), nil
	case ev.BoolValue != nil:
		return strconv.FormatBool(*ev.BoolValue), nil
	case ev.FormulaValue != nil:
		return *ev.FormulaValue, nil
	}
	return "", nil
}

func (g *gsheetTable) MergeCells(startCol, startRow, endCol, endRow int) error {
	if endCol < startCol || endRow < startRow {
		return fmt.Errorf("invalid merge range")
	}
	g.merges = append(g.merges, g.gridRange(startCol, startRow, endCol, endRow))
	return nil
}

func (g *gsheetTable) IsCellMerged(col, row int) bool {
	for _, m := range g.merges {
		if inRange(m, col, row) {
			return true
		}
	}
	return false
}

func (g *gsheetTable) IsCellMergedHorizontally(col, row int) bool {
	for _, m := range g.merges {
		if inRange(m, col, row) && (m.EndColumnIndex-m.StartColumnIndex) > 1 && (m.EndRowIndex-m.StartRowIndex) == 1 {
			return true
		}
	}
	return false
}

func inRange(m *sheets.GridRange, col, row int) bool {
	c, r := int64(col-1), int64(row-1)
	return r >= m.StartRowIndex && r < m.EndRowIndex && c >= m.StartColumnIndex && c < m.EndColumnIndex
}

func (g *gsheetTable) ApplyStyleToCell(col, row int, style spit.Style) error {
	applyStyle(g.format(col, row), style)
	return nil
}

func (g *gsheetTable) ApplyStyleToRange(startCol, startRow, endCol, endRow int, style spit.Style) error {
	for row := startRow; row <= endRow; row++ {
		for col := startCol; col <= endCol; col++ {
			applyStyle(g.format(col, row), style)
		}
	}
	return nil
}

func (g *gsheetTable) ApplyBorderToCell(col, row int, side string, border *spit.Border) error {
	if border == nil || border.Style == spit.BorderStyleNone {
		return nil
	}
	cf := g.format(col, row)
	if cf.Borders == nil {
		cf.Borders = &sheets.Borders{}
	}
	b := &sheets.Border{Style: borderStyle(border.Style), Color: blackColor()}
	switch side {
	case "left":
		cf.Borders.Left = b
	case "right":
		cf.Borders.Right = b
	case "top":
		cf.Borders.Top = b
	case "bottom":
		cf.Borders.Bottom = b
	default:
		return fmt.Errorf("unsupported border side: %s", side)
	}
	return nil
}

func (g *gsheetTable) ApplyBordersToRange(startCol, startRow, endCol, endRow int, borders spit.Borders) error {
	for row := startRow; row <= endRow; row++ {
		for col := startCol; col <= endCol; col++ {
			if col == startCol {
				if err := g.ApplyBorderToCell(col, row, "left", borders.Left); err != nil {
					return err
				}
			}
			if col == endCol {
				if err := g.ApplyBorderToCell(col, row, "right", borders.Right); err != nil {
					return err
				}
			}
			if row == startRow {
				if err := g.ApplyBorderToCell(col, row, "top", borders.Top); err != nil {
					return err
				}
			}
			if row == endRow {
				if err := g.ApplyBorderToCell(col, row, "bottom", borders.Bottom); err != nil {
					return err
				}
			}
		}
	}
	return nil
}

func (g *gsheetTable) HasExistingBorder(col, row int, side string) bool {
	c := g.peek(col, row)
	if c == nil || c.UserEnteredFormat == nil || c.UserEnteredFormat.Borders == nil {
		return false
	}
	b := c.UserEnteredFormat.Borders
	switch side {
	case "left":
		return b.Left != nil
	case "right":
		return b.Right != nil
	case "top":
		return b.Top != nil
	case "bottom":
		return b.Bottom != nil
	}
	return false
}

func (g *gsheetTable) GetColumnLetter(col int) string { return columnLetter(col) }

func (g *gsheetTable) ProcessValue(value interface{}, format string) (interface{}, error) {
	switch v := value.(type) {
	case []interface{}:
		if g.table.ListSeparator != "" {
			return spit.ConvertSliceToString(v, format, g.table.ListSeparator)
		}
		return fmt.Sprintf("%v", v), nil
	default:
		switch format {
		case spit.ExcelizeFormatDefault, spit.ExcelizeFormatFormula, spit.ExcelizeFormatHyperlink:
			return fmt.Sprintf("%v", value), nil
		default:
			if format != "" {
				formatted, err := spit.FormatValue(value, format)
				if err != nil {
					return "", err
				}
				return fmt.Sprintf("%v", formatted), nil
			}
			return fmt.Sprintf("%v", value), nil
		}
	}
}

func (g *gsheetTable) SetCellFormula(col, row int, formula string) error {
	f := formula
	g.cell(col, row).UserEnteredValue = &sheets.ExtendedValue{FormulaValue: &f}
	return nil
}

func (g *gsheetTable) SetCellHyperLink(col, row int, link string) error {
	formula := fmt.Sprintf("=HYPERLINK(%q)", link)
	g.cell(col, row).UserEnteredValue = &sheets.ExtendedValue{FormulaValue: &formula}
	return nil
}

// SetCellImage renders a URL image via the =IMAGE() formula. Embedded bytes cannot be
// placed in a Sheets cell, so they fall back to the image's alt text.
func (g *gsheetTable) SetCellImage(col, row int, img spit.Image) error {
	if img.URL != "" {
		formula := fmt.Sprintf("=IMAGE(%q)", img.URL)
		g.cell(col, row).UserEnteredValue = &sheets.ExtendedValue{FormulaValue: &formula}
		return nil
	}
	return g.SetCellValue(col, row, img.AltText)
}
