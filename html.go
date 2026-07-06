// html.go - HTML export logic for go-spit
//
// This file provides functions to write tabular data to HTML files.
//
// The HTML backend implements the TableOperations interface on top of an in-memory
// cell grid. This lets it reuse the exact same merging (ProcessMerging) and styling
// (RenderStyles) pipelines as the XLSX backend: merges become rowspan/colspan and
// Style values become inline CSS. Document-level presentation (title, description,
// page font/background, custom CSS) is configured through HTMLOptions.

package spit

import (
	"fmt"
	"html"
	"io"
	"strings"
)

// HTMLOptions configures document-level presentation for HTML exports.
// Cell-, column-, row- and header-level styling is driven by the Style values on the
// Table itself (see Column.WithStyle, RowOptions.WithStyle, HeaderOptions, etc.), so
// HTMLOptions only carries options that have no equivalent in the tabular model.
type HTMLOptions struct {
	Title           string    // Optional document title, rendered as an <h1> above the table (and as <title> in full documents)
	Description     string    // Optional description, rendered as a <p> below the title
	BodyStyle       *Style    // Optional default style applied to the page body (font family, background, text color, etc.)
	TableStyle      *Style    // Optional style applied to the <table> element itself
	CustomCSS       string    // Optional raw CSS injected into a <style> block for advanced customization
	Lang            string    // Optional lang attribute for the <html> element (default: "en")
	FragmentOnly    bool      // When true, emit only the title/description/table markup without <!DOCTYPE>, <html>, <head> and <body> wrappers
	Theme           HTMLTheme // Optional built-in stylesheet applied for a polished default look (default: none)
	TableOfContents bool      // When true (documents only), render a linked table of contents from the document headings
}

// HTMLTheme selects a built-in stylesheet injected into the document.
type HTMLTheme int

const (
	// HTMLThemeNone applies no built-in stylesheet (default). Only explicit styles are used.
	HTMLThemeNone HTMLTheme = iota

	// HTMLThemeDefault injects a clean, readable stylesheet: a modern font stack, comfortable
	// spacing, styled headings, zebra-striped tables and a responsive max width.
	HTMLThemeDefault
)

// ExportHTML writes table data to an HTML file using the generic file writer.
// The table is rendered as an HTML <table>, reusing the shared merging and styling
// pipelines. Document-level presentation is controlled through opts.
func ExportHTML(t *Table, opts HTMLOptions, params FileWriteParams) (*FileWriteResult, error) {
	if t == nil {
		return nil, fmt.Errorf("no table provided")
	}

	// Ensure Extension is set for HTML files
	if params.Extension == "" {
		params.Extension = FormatHTML.String()
	}

	L().Info("Starting HTML export to file", String("filename", params.Filename))

	export := &htmlExport{
		table: t,
		opts:  opts,
		grid:  make(map[int]map[int]*htmlCell),
	}

	// Populate the in-memory grid and apply merging/styling via the shared pipelines.
	if err := export.build(); err != nil {
		L().Error("Failed to build HTML table", Error(err))
		return nil, err
	}

	markup := export.render()

	writeFunc := func(writer io.Writer) error {
		_, err := io.WriteString(writer, markup)
		return err
	}

	result, err := params.WriteToFile(writeFunc)
	if err != nil {
		L().Error("Failed to write HTML to file", Error(err))
		return nil, err
	}

	L().Info("HTML export completed", String("filename", params.Filename))
	return result, nil
}

// htmlCell represents a single cell in the in-memory HTML grid.
type htmlCell struct {
	value   string  // Display text (already processed/formatted)
	link    string  // External hyperlink URL; when set the value is wrapped in an <a> tag
	image   *Image  // When set, the cell renders an <img> instead of text
	style   *Style  // Accumulated style for this cell
	borders Borders // Per-side border configuration
	colspan int     // Horizontal span (1 = no span); set on a merge origin
	rowspan int     // Vertical span (1 = no span); set on a merge origin
	covered bool    // True when this cell is absorbed by a merge origin and must not be rendered
	numeric bool    // True when the source value was numeric (used for automatic right alignment)
}

// htmlExport implements TableOperations on top of an in-memory cell grid and
// serializes the result to HTML markup.
type htmlExport struct {
	table   *Table
	opts    HTMLOptions
	caption string                    // Optional <caption> text (set by table blocks)
	grid    map[int]map[int]*htmlCell // grid[row][col], both 1-based
	maxRow  int
	maxCol  int
}

// build populates the grid from the table (preamble, headers, data) and then applies
// the shared merging and styling pipelines, mirroring the XLSX write flow.
func (h *htmlExport) build() error {
	t := h.table
	currentRow := 1

	if len(t.Preamble) > 0 {
		for i, row := range t.Preamble {
			for j, val := range row.Values {
				if err := h.SetCellValue(j+1, currentRow+i, val); err != nil {
					return fmt.Errorf("failed to write preamble cell at (%d, %d): %w", j+1, currentRow+i, err)
				}
			}
		}
		currentRow += len(t.Preamble)
	}

	if t.WriteHeader && len(t.Columns) > 0 {
		headerRows, err := h.writeHeaders(currentRow)
		if err != nil {
			return fmt.Errorf("failed to write headers: %w", err)
		}
		currentRow += headerRows
	}

	flatColumns := t.Columns.GetFlattenedColumns()
	for _, item := range t.Data {
		colIndex := 1
		for _, column := range flatColumns {
			if err := h.writeCell(item, column, colIndex, currentRow); err != nil {
				return fmt.Errorf("failed to write cell: %w", err)
			}
			colIndex++
		}
		currentRow++
	}

	if err := t.ProcessMerging(h); err != nil {
		return fmt.Errorf("failed to process merging: %w", err)
	}

	if err := t.RenderStyles(h); err != nil {
		return fmt.Errorf("failed to render styles: %w", err)
	}

	return nil
}

// writeHeaders writes multi-level header labels starting at startRow.
// Returns the number of header rows written.
func (h *htmlExport) writeHeaders(startRow int) (int, error) {
	t := h.table
	maxDepth := t.Columns.GetMaxDepth()
	if maxDepth == 1 {
		for i, column := range t.Columns {
			if err := h.SetCellValue(i+1, startRow, column.Label); err != nil {
				return 0, fmt.Errorf("failed to set header cell value for column %s: %w", column.Name, err)
			}
		}
		return 1, nil
	}

	maxRow := startRow + maxDepth - 1
	if err := h.writeHeaderRow(t.Columns, startRow, maxRow, 1); err != nil {
		return 0, err
	}
	return maxDepth, nil
}

// writeHeaderRow recursively writes header labels for hierarchical columns.
func (h *htmlExport) writeHeaderRow(columns Columns, currentRow, maxRow, startCol int) error {
	currentCol := startCol
	for _, column := range columns {
		if err := h.SetCellValue(currentCol, currentRow, column.Label); err != nil {
			return fmt.Errorf("failed to set header cell value for column %s at (%d, %d): %w", column.Name, currentCol, currentRow, err)
		}
		if column.HasSubColumns() {
			if currentRow < maxRow {
				if err := h.writeHeaderRow(column.Columns, currentRow+1, maxRow, currentCol); err != nil {
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

// writeCell writes a single data cell, looking up and formatting its value.
// The hyperlink format renders the value as a clickable <a> element.
func (h *htmlExport) writeCell(item Data, column *Column, colIndex, rowIndex int) error {
	value, err, found := item.Lookup(column.Name)
	if err == nil && !found {
		return nil
	}
	if err != nil {
		return fmt.Errorf("error looking up value for column %s: %w", column.Name, err)
	}

	// Image values render as an <img> element rather than text.
	if img, ok := asImage(value); ok {
		return h.SetCellImage(colIndex, rowIndex, img)
	}

	processedValue, err := h.ProcessValue(value, column.Format)
	if err != nil {
		return fmt.Errorf("error processing value for column %s: %w", column.Name, err)
	}

	text := fmt.Sprintf("%v", processedValue)
	if err := h.SetCellValue(colIndex, rowIndex, text); err != nil {
		return fmt.Errorf("error setting cell value for column %s at (%d, %d): %w", column.Name, colIndex, rowIndex, err)
	}

	// Remember numeric source values so they can be right-aligned automatically.
	if column.Format == "" && isNumericValue(value) {
		h.cell(colIndex, rowIndex).numeric = true
	}

	if column.Format == ExcelizeFormatHyperlink {
		if err := h.SetCellHyperLink(colIndex, rowIndex, text); err != nil {
			return fmt.Errorf("error setting hyperlink for column %s at (%d, %d): %w", column.Name, colIndex, rowIndex, err)
		}
	}
	return nil
}

// ---- Grid helpers -----------------------------------------------------------

// cell returns the cell at (col, row), creating it (and expanding the grid bounds) if needed.
func (h *htmlExport) cell(col, row int) *htmlCell {
	if h.grid[row] == nil {
		h.grid[row] = make(map[int]*htmlCell)
	}
	c := h.grid[row][col]
	if c == nil {
		c = &htmlCell{colspan: 1, rowspan: 1}
		h.grid[row][col] = c
	}
	if row > h.maxRow {
		h.maxRow = row
	}
	if col > h.maxCol {
		h.maxCol = col
	}
	return c
}

// peek returns the cell at (col, row) without creating it (nil if absent).
func (h *htmlExport) peek(col, row int) *htmlCell {
	if h.grid[row] == nil {
		return nil
	}
	return h.grid[row][col]
}

// ---- TableOperations implementation ----------------------------------------

// GetTable returns the underlying Table struct.
func (h *htmlExport) GetTable() *Table { return h.table }

// GetCellValue returns the display value of a cell (empty string if absent).
func (h *htmlExport) GetCellValue(col, row int) (string, error) {
	if c := h.peek(col, row); c != nil {
		return c.value, nil
	}
	return "", nil
}

// SetCellValue sets the display value of a cell.
func (h *htmlExport) SetCellValue(col, row int, value interface{}) error {
	h.cell(col, row).value = fmt.Sprintf("%v", value)
	return nil
}

// MergeCells records a rectangular merge: the top-left cell becomes the origin
// (carrying the span) and every other cell in the range is marked as covered.
func (h *htmlExport) MergeCells(startCol, startRow, endCol, endRow int) error {
	if endCol < startCol || endRow < startRow {
		return fmt.Errorf("invalid merge range")
	}
	origin := h.cell(startCol, startRow)
	origin.colspan = endCol - startCol + 1
	origin.rowspan = endRow - startRow + 1
	for row := startRow; row <= endRow; row++ {
		for col := startCol; col <= endCol; col++ {
			if col == startCol && row == startRow {
				continue
			}
			h.cell(col, row).covered = true
		}
	}
	return nil
}

// IsCellMerged reports whether the cell participates in any merge (origin or covered).
func (h *htmlExport) IsCellMerged(col, row int) bool {
	c := h.peek(col, row)
	return c != nil && (c.covered || c.colspan > 1 || c.rowspan > 1)
}

// IsCellMergedHorizontally reports whether the cell is a purely horizontal merge origin.
func (h *htmlExport) IsCellMergedHorizontally(col, row int) bool {
	c := h.peek(col, row)
	return c != nil && c.colspan > 1 && c.rowspan <= 1
}

// ApplyBorderToCell applies a border to one side of a cell.
func (h *htmlExport) ApplyBorderToCell(col, row int, side string, border *Border) error {
	if border == nil || border.Style == BorderStyleNone {
		return nil
	}
	c := h.cell(col, row)
	switch side {
	case "left":
		c.borders.Left = border
	case "right":
		c.borders.Right = border
	case "top":
		c.borders.Top = border
	case "bottom":
		c.borders.Bottom = border
	default:
		return fmt.Errorf("unsupported border side: %s", side)
	}
	return nil
}

// ApplyBordersToRange applies edge borders to the outer cells of a range.
func (h *htmlExport) ApplyBordersToRange(startCol, startRow, endCol, endRow int, borders Borders) error {
	for row := startRow; row <= endRow; row++ {
		for col := startCol; col <= endCol; col++ {
			if col == startCol {
				if err := h.ApplyBorderToCell(col, row, "left", borders.Left); err != nil {
					return err
				}
			}
			if col == endCol {
				if err := h.ApplyBorderToCell(col, row, "right", borders.Right); err != nil {
					return err
				}
			}
			if row == startRow {
				if err := h.ApplyBorderToCell(col, row, "top", borders.Top); err != nil {
					return err
				}
			}
			if row == endRow {
				if err := h.ApplyBorderToCell(col, row, "bottom", borders.Bottom); err != nil {
					return err
				}
			}
		}
	}
	return nil
}

// HasExistingBorder reports whether a border is already set on the given side.
func (h *htmlExport) HasExistingBorder(col, row int, side string) bool {
	c := h.peek(col, row)
	if c == nil {
		return false
	}
	switch side {
	case "left":
		return borderSet(c.borders.Left)
	case "right":
		return borderSet(c.borders.Right)
	case "top":
		return borderSet(c.borders.Top)
	case "bottom":
		return borderSet(c.borders.Bottom)
	}
	return false
}

// ApplyStyleToCell overlays a style onto a cell, preserving previously set properties.
func (h *htmlExport) ApplyStyleToCell(col, row int, style Style) error {
	mergeStyleInto(h.cell(col, row), style)
	return nil
}

// ApplyStyleToRange overlays a style onto every cell in a range.
func (h *htmlExport) ApplyStyleToRange(startCol, startRow, endCol, endRow int, style Style) error {
	for row := startRow; row <= endRow; row++ {
		for col := startCol; col <= endCol; col++ {
			mergeStyleInto(h.cell(col, row), style)
		}
	}
	return nil
}

// GetColumnLetter returns the spreadsheet-style column letter for a 1-based index.
func (h *htmlExport) GetColumnLetter(col int) string {
	if col <= 0 {
		return ""
	}
	var sb strings.Builder
	for col > 0 {
		col--
		sb.WriteByte(byte('A' + col%26))
		col /= 26
	}
	// Reverse the accumulated letters.
	b := []byte(sb.String())
	for i, j := 0, len(b)-1; i < j; i, j = i+1, j-1 {
		b[i], b[j] = b[j], b[i]
	}
	return string(b)
}

// ProcessValue formats a value for output and merge comparison, mirroring the
// semantics used by the other backends so merge decisions stay consistent.
func (h *htmlExport) ProcessValue(value interface{}, format string) (interface{}, error) {
	if img, ok := asImage(value); ok {
		return img.TextValue(), nil
	}
	switch v := value.(type) {
	case []interface{}:
		if h.table.ListSeparator != "" {
			return ConvertSliceToString(v, format, h.table.ListSeparator)
		}
		return fmt.Sprintf("%v", v), nil
	default:
		switch format {
		case ExcelizeFormatDefault, ExcelizeFormatFormula, ExcelizeFormatHyperlink:
			return fmt.Sprintf("%v", value), nil
		default:
			if format != "" {
				formatted, err := FormatValue(value, format)
				if err != nil {
					return "", err
				}
				value = formatted
			}
			return fmt.Sprintf("%v", value), nil
		}
	}
}

// SetCellFormula stores the formula text as the cell's display value (HTML has no formulas).
func (h *htmlExport) SetCellFormula(col, row int, formula string) error {
	h.cell(col, row).value = formula
	return nil
}

// SetCellHyperLink marks a cell as a hyperlink so it is rendered as an <a> element.
func (h *htmlExport) SetCellHyperLink(col, row int, link string) error {
	h.cell(col, row).link = link
	return nil
}

// SetCellImage marks a cell so it renders an <img> element.
func (h *htmlExport) SetCellImage(col, row int, img Image) error {
	imgCopy := img
	h.cell(col, row).image = &imgCopy
	return nil
}

// ---- Serialization ----------------------------------------------------------

// render serializes a single-table export to full HTML markup.
func (h *htmlExport) render() string {
	var b strings.Builder
	writeDocumentOpen(&b, h.opts)
	h.writeTable(&b)
	writeDocumentClose(&b, h.opts)
	return b.String()
}

// writeDocumentOpen writes the document preamble: the optional <!DOCTYPE>/<head>/<body>
// wrappers (unless FragmentOnly), followed by the visible title and description.
func writeDocumentOpen(b *strings.Builder, opts HTMLOptions) {
	if !opts.FragmentOnly {
		lang := opts.Lang
		if lang == "" {
			lang = "en"
		}
		b.WriteString("<!DOCTYPE html>\n")
		b.WriteString(fmt.Sprintf("<html lang=\"%s\">\n", html.EscapeString(lang)))
		b.WriteString("<head>\n<meta charset=\"utf-8\">\n")
		b.WriteString("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n")
		if opts.Title != "" {
			b.WriteString(fmt.Sprintf("<title>%s</title>\n", html.EscapeString(opts.Title)))
		}
		// Built-in theme first, so CustomCSS can override it.
		if css := themeCSS(opts.Theme); css != "" {
			b.WriteString("<style>" + css + "</style>\n")
		}
		if opts.CustomCSS != "" {
			b.WriteString(fmt.Sprintf("<style>%s</style>\n", opts.CustomCSS))
		}
		b.WriteString("</head>\n")
		if css := styleToCSS(opts.BodyStyle); css != "" {
			b.WriteString(fmt.Sprintf("<body style=\"%s\">\n", css))
		} else {
			b.WriteString("<body>\n")
		}
	} else {
		// Fragment mode has no <head>; emit any stylesheet as a leading <style> block.
		if css := themeCSS(opts.Theme); css != "" {
			b.WriteString("<style>" + css + "</style>\n")
		}
		if opts.CustomCSS != "" {
			b.WriteString(fmt.Sprintf("<style>%s</style>\n", opts.CustomCSS))
		}
	}

	if opts.Title != "" {
		b.WriteString(fmt.Sprintf("<h1>%s</h1>\n", html.EscapeString(opts.Title)))
	}
	if opts.Description != "" {
		b.WriteString(fmt.Sprintf("<p>%s</p>\n", html.EscapeString(opts.Description)))
	}
}

// writeDocumentClose closes the document wrappers opened by writeDocumentOpen.
func writeDocumentClose(b *strings.Builder, opts HTMLOptions) {
	if !opts.FragmentOnly {
		b.WriteString("</body>\n</html>\n")
	}
}

// writeTable serializes the cell grid as a standalone <table> element.
func (h *htmlExport) writeTable(b *strings.Builder) {
	t := h.table
	opts := h.opts

	tableStyle := "border-collapse:collapse"
	if css := styleToCSS(opts.TableStyle); css != "" {
		tableStyle += ";" + css
	}
	b.WriteString(fmt.Sprintf("<table style=\"%s\">\n", tableStyle))

	if h.caption != "" {
		b.WriteString(fmt.Sprintf("<caption>%s</caption>\n", html.EscapeString(h.caption)))
	}

	h.writeColgroup(b)

	// Determine which rows are header rows (rendered with <th>).
	headerStart := t.GetHeaderStartRow()
	headerEnd := headerStart - 1 // no header rows by default
	if t.WriteHeader && len(t.Columns) > 0 {
		headerEnd = headerStart + t.Columns.GetMaxDepth() - 1
	}

	// The <thead> spans every row above the data (preamble rows and header rows);
	// the <tbody> holds the data rows.
	theadEnd := headerStart - 1 // preamble rows only
	if headerEnd >= headerStart {
		theadEnd = headerEnd
	}

	if theadEnd >= 1 {
		b.WriteString("<thead>\n")
		for row := 1; row <= theadEnd; row++ {
			h.writeRow(b, row, headerStart, headerEnd)
		}
		b.WriteString("</thead>\n")
	}
	if h.maxRow > theadEnd {
		b.WriteString("<tbody>\n")
		for row := theadEnd + 1; row <= h.maxRow; row++ {
			h.writeRow(b, row, headerStart, headerEnd)
		}
		b.WriteString("</tbody>\n")
	}

	b.WriteString("</table>\n")
}

// writeRow serializes a single grid row, skipping cells absorbed by a merge.
func (h *htmlExport) writeRow(b *strings.Builder, row, headerStart, headerEnd int) {
	b.WriteString("<tr>\n")
	isHeader := row >= headerStart && row <= headerEnd
	for col := 1; col <= h.maxCol; col++ {
		c := h.peek(col, row)
		if c != nil && c.covered {
			continue
		}
		h.renderCell(b, c, col, row, isHeader)
	}
	b.WriteString("</tr>\n")
}

// writeColgroup emits a <colgroup> mapping each leaf column's Width (in character units)
// to a CSS ch width. It is skipped entirely when no column has an explicit width.
func (h *htmlExport) writeColgroup(b *strings.Builder) {
	flat := h.table.Columns.GetFlattenedColumns()
	hasWidth := false
	for _, c := range flat {
		if c.Width > 0 {
			hasWidth = true
			break
		}
	}
	if !hasWidth {
		return
	}
	b.WriteString("<colgroup>\n")
	for _, c := range flat {
		if c.Width > 0 {
			b.WriteString(fmt.Sprintf("<col style=\"width:%gch\">\n", c.Width))
		} else {
			b.WriteString("<col>\n")
		}
	}
	b.WriteString("</colgroup>\n")
}

// renderCell serializes a single (non-covered) cell as a <td> or <th> element.
func (h *htmlExport) renderCell(b *strings.Builder, c *htmlCell, col, row int, isHeader bool) {
	tag := "td"
	if isHeader {
		tag = "th"
	}

	colspan, rowspan := 1, 1
	text, link := "", ""
	var image *Image
	var style *Style
	var borders Borders
	if c != nil {
		colspan = max(c.colspan, 1)
		rowspan = max(c.rowspan, 1)
		text = c.value
		link = c.link
		image = c.image
		style = c.style
		borders = h.effectiveBorders(col, row, colspan, rowspan)
	}

	// The theme's stylesheet controls cell padding; otherwise apply a small inline default.
	basePadding := "padding:4px 8px"
	if h.opts.Theme != HTMLThemeNone {
		basePadding = ""
	}

	// Right-align numeric data cells that carry no explicit alignment.
	numericAlign := ""
	if c != nil && c.numeric && (style == nil || style.Alignment == AlignmentNone) {
		numericAlign = "text-align:right"
	}

	css := combineCSS(basePadding, styleToCSS(style), numericAlign, bordersToCSS(borders))

	var attrs strings.Builder
	if colspan > 1 {
		attrs.WriteString(fmt.Sprintf(" colspan=\"%d\"", colspan))
	}
	if rowspan > 1 {
		attrs.WriteString(fmt.Sprintf(" rowspan=\"%d\"", rowspan))
	}
	if isHeader {
		attrs.WriteString(" scope=\"col\"")
	}

	var content string
	if image != nil {
		content = imgTag(*image)
	} else {
		content = html.EscapeString(text)
	}
	if link != "" {
		content = fmt.Sprintf("<a href=\"%s\">%s</a>", html.EscapeString(link), content)
	}

	styleAttr := ""
	if css != "" {
		styleAttr = fmt.Sprintf(" style=\"%s\"", css)
	}
	b.WriteString(fmt.Sprintf("<%s%s%s>%s</%s>\n", tag, attrs.String(), styleAttr, content, tag))
}

// effectiveBorders computes the outer borders of a (possibly spanned) cell by
// scanning the edge cells of its merge range. This preserves borders that the
// styling pipeline applied to cells later absorbed by a merge.
func (h *htmlExport) effectiveBorders(col, row, colspan, rowspan int) Borders {
	var res Borders
	lastCol := col + colspan - 1
	lastRow := row + rowspan - 1

	for c := col; c <= lastCol; c++ {
		if cell := h.peek(c, row); cell != nil && borderSet(cell.borders.Top) {
			res.Top = cell.borders.Top
			break
		}
	}
	for c := col; c <= lastCol; c++ {
		if cell := h.peek(c, lastRow); cell != nil && borderSet(cell.borders.Bottom) {
			res.Bottom = cell.borders.Bottom
			break
		}
	}
	for r := row; r <= lastRow; r++ {
		if cell := h.peek(col, r); cell != nil && borderSet(cell.borders.Left) {
			res.Left = cell.borders.Left
			break
		}
	}
	for r := row; r <= lastRow; r++ {
		if cell := h.peek(lastCol, r); cell != nil && borderSet(cell.borders.Right) {
			res.Right = cell.borders.Right
			break
		}
	}
	return res
}

// ---- CSS helpers ------------------------------------------------------------

// isNumericValue reports whether a raw cell value is a numeric type (used to right-align it).
func isNumericValue(value interface{}) bool {
	switch value.(type) {
	case int, int8, int16, int32, int64,
		uint, uint8, uint16, uint32, uint64,
		float32, float64:
		return true
	}
	return false
}

// themeCSS returns the built-in stylesheet for the given theme, or "" for HTMLThemeNone.
func themeCSS(theme HTMLTheme) string {
	switch theme {
	case HTMLThemeDefault:
		return "" +
			"body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;" +
			"line-height:1.5;color:#24292f;max-width:960px;margin:2rem auto;padding:0 1rem;}" +
			"h1,h2,h3,h4,h5,h6{line-height:1.25;margin:1.4em 0 .5em;}" +
			"h1{font-size:2em;border-bottom:1px solid #eaecef;padding-bottom:.3em;}" +
			"h2{font-size:1.5em;border-bottom:1px solid #eaecef;padding-bottom:.3em;}" +
			"p{margin:.6em 0;}" +
			"a{color:#0969da;text-decoration:none;}a:hover{text-decoration:underline;}" +
			"ul,ol{margin:.6em 0;padding-left:1.6em;}li{margin:.2em 0;}" +
			"table{border-collapse:collapse;width:100%;margin:1em 0;overflow:auto;display:block;}" +
			"caption{caption-side:top;font-weight:600;text-align:left;margin-bottom:.4em;}" +
			"th,td{border:1px solid #d0d7de;padding:6px 13px;}" +
			"thead th{background:#f6f8fa;}" +
			"tbody tr:nth-child(even){background:#f6f8fa;}" +
			"blockquote{margin:.8em 0;padding:0 1em;color:#57606a;border-left:.25em solid #d0d7de;}" +
			"pre{background:#f6f8fa;padding:1em;overflow:auto;border-radius:6px;}" +
			"code{font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;font-size:.9em;}" +
			"pre code{background:none;padding:0;}" +
			"dl{margin:.6em 0;}dt{font-weight:600;}dd{margin:0 0 .4em 1em;color:#57606a;}" +
			"hr{border:0;border-top:1px solid #d0d7de;margin:1.5em 0;}" +
			"nav.toc{background:#f6f8fa;border:1px solid #d0d7de;border-radius:6px;padding:.5em 1em;margin:1em 0;}" +
			"nav.toc ul{list-style:none;padding-left:0;}"
	default:
		return ""
	}
}

// imgTag builds an <img> element for a cell image. Embedded content is emitted as a
// base64 data URI; otherwise the URL is used as the source. Empty images render nothing.
func imgTag(img Image) string {
	src := img.URL
	if img.HasBytes() {
		src = img.DataURI()
	}
	if src == "" {
		return html.EscapeString(img.AltText)
	}
	var attrs strings.Builder
	attrs.WriteString(fmt.Sprintf("<img src=\"%s\" alt=\"%s\"", html.EscapeString(src), html.EscapeString(img.AltText)))
	if img.Width > 0 {
		attrs.WriteString(fmt.Sprintf(" width=\"%d\"", img.Width))
	}
	if img.Height > 0 {
		attrs.WriteString(fmt.Sprintf(" height=\"%d\"", img.Height))
	}
	attrs.WriteString(">")
	return attrs.String()
}

// mergeStyleInto overlays the set fields of style onto the cell's existing style.
func mergeStyleInto(c *htmlCell, style Style) {
	if c.style == nil {
		cp := style
		c.style = &cp
		return
	}
	cur := c.style
	if style.Bold {
		cur.Bold = true
	}
	if style.Italic {
		cur.Italic = true
	}
	if style.Underline != "" {
		cur.Underline = style.Underline
	}
	if style.TextColor != "" {
		cur.TextColor = style.TextColor
	}
	if style.BackgroundColor != "" {
		cur.BackgroundColor = style.BackgroundColor
	}
	if style.FontSize > 0 {
		cur.FontSize = style.FontSize
	}
	if style.FontFamily != "" {
		cur.FontFamily = style.FontFamily
	}
	if style.Alignment != AlignmentNone {
		cur.Alignment = style.Alignment
	}
	if style.NumFmt != "" {
		cur.NumFmt = style.NumFmt
	}
}

// styleToCSS converts a Style to an inline CSS declaration string (empty if nil/blank).
func styleToCSS(s *Style) string {
	if s == nil {
		return ""
	}
	var parts []string
	if s.Bold {
		parts = append(parts, "font-weight:bold")
	}
	if s.Italic {
		parts = append(parts, "font-style:italic")
	}
	if s.Underline != "" {
		parts = append(parts, "text-decoration:underline")
	}
	if s.TextColor != "" {
		parts = append(parts, "color:"+cssColor(s.TextColor))
	}
	if s.BackgroundColor != "" {
		parts = append(parts, "background-color:"+cssColor(s.BackgroundColor))
	}
	if s.FontSize > 0 {
		parts = append(parts, fmt.Sprintf("font-size:%gpt", s.FontSize))
	}
	if s.FontFamily != "" {
		parts = append(parts, "font-family:"+cssFontFamily(s.FontFamily))
	}
	if s.Alignment != AlignmentNone {
		horizontal, vertical := s.Alignment.GetAlignmentValues()
		parts = append(parts, "text-align:"+horizontal)
		parts = append(parts, "vertical-align:"+cssVerticalAlign(vertical))
	}
	return strings.Join(parts, ";")
}

// bordersToCSS converts a Borders configuration to inline CSS border declarations.
func bordersToCSS(b Borders) string {
	var parts []string
	if borderSet(b.Left) {
		parts = append(parts, "border-left:"+borderStyleToCSS(b.Left.Style))
	}
	if borderSet(b.Right) {
		parts = append(parts, "border-right:"+borderStyleToCSS(b.Right.Style))
	}
	if borderSet(b.Top) {
		parts = append(parts, "border-top:"+borderStyleToCSS(b.Top.Style))
	}
	if borderSet(b.Bottom) {
		parts = append(parts, "border-bottom:"+borderStyleToCSS(b.Bottom.Style))
	}
	return strings.Join(parts, ";")
}

// borderStyleToCSS maps a BorderStyle to a CSS border shorthand value.
func borderStyleToCSS(style BorderStyle) string {
	switch style {
	case BorderStyleThin:
		return "1px solid #000000"
	case BorderStyleMedium:
		return "2px solid #000000"
	case BorderStyleThick:
		return "3px solid #000000"
	case BorderStyleDashed:
		return "1px dashed #000000"
	case BorderStyleDotted:
		return "1px dotted #000000"
	case BorderStyleDouble:
		return "3px double #000000"
	default:
		return ""
	}
}

// borderSet reports whether a border is present and visible.
func borderSet(b *Border) bool {
	return b != nil && b.Style != BorderStyleNone
}

// cssColor normalizes a color value for CSS, prefixing bare 6-digit hex codes with '#'.
func cssColor(c string) string {
	if c == "" || c[0] == '#' {
		return c
	}
	if len(c) == 6 && isHex(c) {
		return "#" + c
	}
	return c
}

// isHex reports whether s consists solely of hexadecimal digits.
func isHex(s string) bool {
	for _, r := range s {
		if !((r >= '0' && r <= '9') || (r >= 'a' && r <= 'f') || (r >= 'A' && r <= 'F')) {
			return false
		}
	}
	return true
}

// cssFontFamily quotes a font family name when it contains spaces.
func cssFontFamily(family string) string {
	if strings.ContainsAny(family, " \t") {
		return "'" + family + "'"
	}
	return family
}

// cssVerticalAlign maps an internal vertical alignment token to a CSS value.
func cssVerticalAlign(v string) string {
	if v == "center" {
		return "middle"
	}
	return v
}

// combineCSS joins non-empty CSS declaration fragments with ';'.
func combineCSS(parts ...string) string {
	var nonEmpty []string
	for _, p := range parts {
		if p != "" {
			nonEmpty = append(nonEmpty, p)
		}
	}
	return strings.Join(nonEmpty, ";")
}
