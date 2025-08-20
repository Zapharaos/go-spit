// table.go - Table abstraction and operations.
//
// This file defines the Table structure and the TableOperations interface, which provide a unified way to
// represent, manipulate, and export tabular data with hierarchical columns, formatting, and cell/row configuration.
// It allows for flexible data representation and export across different libraries and backends.
// Implement TableOperations for each backend (e.g., Excelize, CSV) to support table manipulation and export.

//go:generate mockgen -destination=table_mock.go -package=spit . TableOperations

package spit

import (
	"fmt"
	"strings"
)

// TableOperations defines table-specific operations interface.
// Implement this interface to support table manipulation for various libraries and backends.
type TableOperations interface {
	// GetTable returns the underlying Table struct for direct access/manipulation.
	GetTable() *Table

	// GetCellValue returns the value of a cell at the given column and row (1-based indices).
	// The implementation should convert indices to the appropriate cell reference and retrieve the value.
	GetCellValue(col, row int) (string, error)

	// SetCellValue sets the value of a cell at the given column and row.
	// The implementation should convert indices to the appropriate cell reference and set the value.
	SetCellValue(col, row int, value interface{}) error

	// MergeCells merges a rectangular range of cells defined by start and end coordinates.
	MergeCells(startCol, startRow, endCol, endRow int) error

	// IsCellMerged checks if a cell at the given column and row is part of any merged range.
	IsCellMerged(col, row int) bool

	// IsCellMergedHorizontally checks if a cell at the given column and row is merged horizontally (across columns).
	IsCellMergedHorizontally(col, row int) bool

	// ApplyBorderToCell applies a border to a cell on the specified side ("left", "right", "top", "bottom").
	// The border style is defined by the Border parameter.
	ApplyBorderToCell(col, row int, side string, border *Border) error

	// ApplyBordersToRange applies borders to all cells in a rectangular range.
	// Each side of the range can have a different border style, as specified in the Borders parameter.
	ApplyBordersToRange(startCol, startRow, endCol, endRow int, borders Borders) error

	// HasExistingBorder checks if a cell at the given column and row has a border on the specified side.
	HasExistingBorder(col, row int, side string) bool

	// ApplyStyleToCell applies a style to a specific cell at the given column and row.
	ApplyStyleToCell(col, row int, style Style) error

	// ApplyStyleToRange Applies a style to a rectangular range of cells.
	ApplyStyleToRange(startCol, startRow, endCol, endRow int, style Style) error

	// GetColumnLetter Returns the Excel-style column letter (e.g., "A", "B") for a given column index.
	GetColumnLetter(col int) string

	// ProcessValue Processes a value for output, applying formatting if needed.
	ProcessValue(value interface{}, format string) (interface{}, error)
}

// Table represents a structured data table with configuration for export operations.
// Contains data rows, column definitions (including hierarchy and formatting), and options for styling, merging, and headers.
type Table struct {
	Data           DataSlice      // The actual data rows to be exported
	Columns        Columns        // Column definitions including hierarchy and formatting
	RowOptionsMap  RowOptionsMap  // Row-specific options (styling, merging, borders)
	CellOptionsMap CellOptionsMap // Cell-specific options for fine-grained control
	WriteHeader    bool           // Whether to generate headers from column definitions
	Limit          int64          // Maximum number of data rows to export (0 = no limit)
	ListSeparator  string         // separator used when rendering slice/array values as strings
}

// NewTable creates a new Table instance with the provided data slice and column definitions.
func NewTable(slice DataSlice, columns Columns, writeHeader bool) *Table {
	return &Table{
		Data:        slice,
		Columns:     columns,
		WriteHeader: writeHeader,
	}
}

// WithRowOptions allows setting row-specific options for the table.
func (t *Table) WithRowOptions(rowOptions RowOptionsMap) *Table {
	t.RowOptionsMap = rowOptions
	return t
}

// WithCellOptions allows setting cell-specific options for the table.
func (t *Table) WithCellOptions(cellOptions CellOptionsMap) *Table {
	t.CellOptionsMap = cellOptions
	return t
}

// GetDataStartRow calculates the starting row number for data based on header configuration.
// Accounts for multi-level headers by reserving rows for each level of the column hierarchy.
func (t *Table) GetDataStartRow() int {
	dataStartRow := 1
	if t.WriteHeader && len(t.Columns) > 0 {
		// Calculate the maximum depth of the column hierarchy
		// Each level of nesting requires its own header row
		maxDepth := t.Columns.GetMaxDepth()
		dataStartRow = maxDepth + 1
	}
	return dataStartRow
}

// GetDataIndexFromRowIndex converts a 1-based row index to the corresponding 0-based data slice index.
// Adjusts for header rows if present.
func (t *Table) GetDataIndexFromRowIndex(rowIndex int) int {
	if t.WriteHeader {
		// Adjust row index to account for header rows
		headerShift := t.GetDataStartRow()
		return rowIndex - headerShift
	}
	return rowIndex
}

// Data represents a single row of table data as a map from column name to value.
type Data map[string]interface{}

// DataSlice is a slice of Data rows.
type DataSlice []Data

// Lookup recursively looks up a key in a nested map structure.
// Supports multi-level key access for hierarchical data.
func (d Data) Lookup(ks ...string) (rval interface{}, err error, found bool) {
	var ok bool
	if len(ks) == 0 {
		return nil, fmt.Errorf("nestedMapLookup needs at least one key"), false
	}
	if rval, ok = d[strings.TrimSpace(ks[0])]; !ok {
		return nil, nil, false
	} else if len(ks) == 1 {
		return rval, nil, true
	} else if d, ok = rval.(Data); !ok {
		return nil, fmt.Errorf("malformed structure at %#v", rval), false
	} else {
		return d.Lookup(ks[1:]...)
	}
}

// Column represents a single column definition for table exports.
// Columns can be nested to create hierarchical structures, allowing for
// complex header layouts and grouped data organization.
type Column struct {
	Name    string      // Field name in the data source (for leaf columns)
	Label   string      // Display label for headers
	Format  string      // Format specification for value processing (e.g., date format)
	Merge   *MergeRules // Optional merge configuration for this column
	Borders *Borders    // Borders configuration
	Style   *Style      // Optional content style
	Columns Columns     // Sub-columns for hierarchical structures
}

// NewColumn creates a new Column with the specified name and label.
func NewColumn(name, label string) *Column {
	return &Column{
		Name:  name,
		Label: label,
	}
}

// WithFormat sets the format for this column.
func (c *Column) WithFormat(format string) *Column {
	c.Format = format
	return c
}

// WithMerge sets the merge rules for this column.
func (c *Column) WithMerge(merge *MergeRules) *Column {
	c.Merge = merge
	return c
}

// WithBorders sets the borders for this column.
func (c *Column) WithBorders(borders *Borders) *Column {
	c.Borders = borders
	return c
}

// WithStyle sets the style for this column.
func (c *Column) WithStyle(style *Style) *Column {
	c.Style = style
	return c
}

// WithSubColumns sets the sub-columns for this column.
func (c *Column) WithSubColumns(subColumns Columns) *Column {
	c.Columns = subColumns
	return c
}

// AddSubColumn adds a new sub-column to this column.
func (c *Column) AddSubColumn(subColumn *Column) *Column {
	// Initialize sub-columns if not already set
	if c.Columns == nil {
		c.Columns = make(Columns, 0)
	}
	// Append the new sub-column
	c.Columns = append(c.Columns, subColumn)
	return c
}

// RemoveSubColumn removes a sub-column by name from this column.
func (c *Column) RemoveSubColumn(name string) *Column {
	// Filter out the sub-column with the specified name
	newColumns := make(Columns, 0)
	for _, subCol := range c.Columns {
		if subCol.Name != name {
			newColumns = append(newColumns, subCol)
		}
	}
	c.Columns = newColumns
	return c
}

// HasSubColumns returns true if this column contains nested sub-columns.
func (c *Column) HasSubColumns() bool {
	return len(c.Columns) > 0
}

// CountSubColumns returns the total number of leaf columns represented by this column.
// For leaf columns, returns 1. For parent columns, recursively counts all leaf columns in the hierarchy.
func (c *Column) CountSubColumns() int {
	if c.HasSubColumns() {
		// Recursively count all leaf columns in sub-columns
		total := 0
		for _, subCol := range c.Columns {
			total += subCol.CountSubColumns()
		}
		return total
	}
	return 1 // Leaf column represents exactly one data column
}

// Columns represents a collection of column definitions, possibly hierarchical.
type Columns []*Column

// GetTotalColumnCount calculates the total number of leaf columns in the collection.
func (c Columns) GetTotalColumnCount() int {
	total := 0
	for _, column := range c {
		total += column.CountSubColumns()
	}
	return total
}

// GetFlattenedColumns returns a flattened list of all leaf columns in the hierarchy.
// Traverses the entire column structure and extracts only columns without sub-columns.
func (c Columns) GetFlattenedColumns() Columns {
	var flattened Columns
	for _, column := range c {
		if column.HasSubColumns() {
			// Recursively flatten sub-columns
			flattened = append(flattened, column.Columns.GetFlattenedColumns()...)
		} else {
			// Add leaf column directly
			flattened = append(flattened, column)
		}
	}
	return flattened
}

// GetMaxDepth returns the maximum depth of the column hierarchy.
func (c Columns) GetMaxDepth() int {
	maxDepth := 1
	for _, column := range c {
		if column.HasSubColumns() {
			// Recursively calculate depth including this level
			depth := 1 + column.Columns.GetMaxDepth()
			if depth > maxDepth {
				maxDepth = depth
			}
		}
	}
	return maxDepth
}

// RowOptionsMap maps row indices to their specific options.
type RowOptionsMap map[int]RowOptions

// RowOptions represents option settings for a specific row in the export.
// This allows fine-grained control over individual rows, overriding default
// column-based settings when needed.
type RowOptions struct {
	RowIndex  int         // The 0-based index of the row this option applies to
	Border    *Borders    // Optional border configuration for the entire row
	Style     *Style      // Optional style configuration for the entire row
	Merge     *MergeRules // Optional merge configuration that overrides column settings
	Mergeable bool        // Whether this row cells can participate in merge operations
}

// NewRowOptions creates a new RowOptions instance for the specified row index.
func NewRowOptions(rowIndex int) *RowOptions {
	return &RowOptions{
		RowIndex: rowIndex,
	}
}

// WithBorder sets the border configuration for this row.
func (rowOptions *RowOptions) WithBorder(border *Borders) *RowOptions {
	rowOptions.Border = border
	return rowOptions
}

// WithStyle sets the style configuration for this row.
func (rowOptions *RowOptions) WithStyle(style *Style) *RowOptions {
	rowOptions.Style = style
	return rowOptions
}

// WithMerge sets the merge rules for this row, overriding any column-level settings.
func (rowOptions *RowOptions) WithMerge(merge *MergeRules) *RowOptions {
	rowOptions.Merge = merge
	return rowOptions
}

// WithMergeable sets whether this row's cells can participate in external merge operations.
func (rowOptions *RowOptions) WithMergeable(mergeable bool) *RowOptions {
	rowOptions.Mergeable = mergeable
	return rowOptions
}

// CellOptionsMap provides cell-level option mapping.
// The outer map keys are column indices, inner map keys are row indices.
type CellOptionsMap map[int]map[int]CellOptions

// CellOptions represents option settings for a specific cell.
// This provides the finest level of control, allowing individual cells
// to override both column and row settings.
type CellOptions struct {
	RowIndex  int      // The 0-based row index of this cell
	ColIndex  int      // The 0-based column index of this cell
	Border    *Borders // Optional border configuration for this cell
	Style     *Style   // Optional style configuration for this cell
	Mergeable bool     // Whether this cell can participate in merge operations
}

// NewCellOptions creates a new CellOptions instance for the specified row and column indices.
func NewCellOptions(rowIndex, colIndex int) *CellOptions {
	return &CellOptions{
		RowIndex: rowIndex,
		ColIndex: colIndex,
	}
}

// WithBorder sets the border configuration for this cell.
func (cellOptions *CellOptions) WithBorder(border *Borders) *CellOptions {
	cellOptions.Border = border
	return cellOptions
}

// WithStyle sets the style configuration for this cell.
func (cellOptions *CellOptions) WithStyle(style *Style) *CellOptions {
	cellOptions.Style = style
	return cellOptions
}

// WithMergeable sets whether this cell can participate in external merge operations.
func (cellOptions *CellOptions) WithMergeable(mergeable bool) *CellOptions {
	cellOptions.Mergeable = mergeable
	return cellOptions
}

type MergeConditions []MergeCondition

// MergeCondition defines the conditions under which cells should be merged.
// These conditions are evaluated when determining whether adjacent cells
// should be combined into a single merged cell.
type MergeCondition string

const (
	// MergeConditionIdentical merges cells when their values are identical and non-empty
	MergeConditionIdentical MergeCondition = "identical"

	// MergeConditionEmpty merges cells when both values are empty or nil
	MergeConditionEmpty MergeCondition = "empty"
)

// AnyMatch checks if two sets of merge conditions share at least one common condition.
// This is used to determine if two cells or ranges can be merged together based on their configurations.
func (mc MergeConditions) AnyMatch(mc2 []MergeCondition) bool {
	for _, cond1 := range mc {
		for _, cond2 := range mc2 {
			if cond1 == cond2 {
				return true // Found a matching condition
			}
		}
	}
	return false // No compatible conditions found
}

// ValuesShouldMerge determines if two values should be merged based on the specified conditions.
func (mc MergeConditions) ValuesShouldMerge(value1, value2 interface{}) bool {
	if len(mc) == 0 {
		return false // No conditions specified - don't merge
	}

	// Convert values to strings for consistent comparison
	val1Str := strings.TrimSpace(fmt.Sprintf("%v", value1))
	val2Str := strings.TrimSpace(fmt.Sprintf("%v", value2))

	// Determine if values are considered empty
	isEmpty1 := val1Str == "" || val1Str == "<nil>"
	isEmpty2 := val2Str == "" || val2Str == "<nil>"

	// Evaluate each condition to see if any match
	for _, condition := range mc {
		switch condition {
		case MergeConditionIdentical:
			// Merge if values are identical and both are non-empty
			if val1Str == val2Str && !isEmpty1 && !isEmpty2 {
				return true
			}
		case MergeConditionEmpty:
			// Merge if both values are empty
			if isEmpty1 && isEmpty2 {
				return true
			}
		}
	}
	return false // No conditions matched
}

// MergeRules holds merge conditions for columns and rows.
// It defines when and how cells should be merged based on their content.
// Empty conditions arrays mean no merging will be applied.
type MergeRules struct {
	Vertical   MergeConditions `json:"vertical,omitempty"`   // Conditions for merging cells vertically (between rows)
	Horizontal MergeConditions `json:"horizontal,omitempty"` // Conditions for merging cells horizontally (between columns)
}

// NewMergeRules creates a new MergeRules instance with specified vertical and horizontal conditions.
func NewMergeRules(vertical, horizontal MergeConditions) *MergeRules {
	return &MergeRules{
		Vertical:   vertical,
		Horizontal: horizontal,
	}
}

// BorderStyle represents the visual style of entity borders.
// These constants correspond to common border styles available in document applications.
type BorderStyle int

const (
	BorderStyleNone   BorderStyle = iota // No border
	BorderStyleThin                      // Thin solid line
	BorderStyleMedium                    // Medium thickness solid line
	BorderStyleDashed                    // Dashed line
	BorderStyleDotted                    // Dotted line
	BorderStyleThick                     // Thick solid line
	BorderStyleDouble                    // Double line
)

// Border represents the configuration for an entity border.
type Border struct {
	Style BorderStyle // The visual style to apply to this border side
}

func NewBorder(style BorderStyle) *Border {
	return &Border{
		Style: style,
	}
}

// Borders represents all borders configuration for an entity.
type Borders struct {
	Left   *Border  // Left border configuration
	Right  *Border  // Right border configuration
	Top    *Border  // Top border configuration
	Bottom *Border  // Bottom border configuration
	Inner  *Borders // Inner borders for ranges (used in some contexts)
}

// NewBorders creates a Borders with the individual style per edge.
func NewBorders(left, right, top, bottom BorderStyle) *Borders {
	return &Borders{
		Left:   NewBorder(left),
		Right:  NewBorder(right),
		Top:    NewBorder(top),
		Bottom: NewBorder(bottom),
	}
}

// NewBordersBoundaries creates a Borders with the same style applied to all edges.
func NewBordersBoundaries(style BorderStyle) *Borders {
	border := NewBorder(style)
	return &Borders{
		Left:   border,
		Right:  border,
		Top:    border,
		Bottom: border,
	}
}

// HasBorders checks if any borders are configured in these Borders.
func (bc *Borders) HasBorders() bool {
	return (bc.Left != nil && bc.Left.Style != BorderStyleNone) ||
		(bc.Right != nil && bc.Right.Style != BorderStyleNone) ||
		(bc.Top != nil && bc.Top.Style != BorderStyleNone) ||
		(bc.Bottom != nil && bc.Bottom.Style != BorderStyleNone)
}

// SetBoundaries sets all borders (left, right, top, bottom) to the same style.
func (bc *Borders) SetBoundaries(style BorderStyle) *Borders {
	bc.SetVertical(style)
	bc.SetHorizontal(style)
	return bc
}

// SetVertical sets both left and right borders to the same style.
func (bc *Borders) SetVertical(style BorderStyle) *Borders {
	bc.SetLeft(style)
	bc.SetRight(style)
	return bc
}

// SetHorizontal sets both top and bottom borders to the same style.
func (bc *Borders) SetHorizontal(style BorderStyle) *Borders {
	bc.SetTop(style)
	bc.SetBottom(style)
	return bc
}

// SetLeft sets the left border style.
func (bc *Borders) SetLeft(style BorderStyle) *Borders {
	if bc.Left == nil {
		bc.Left = NewBorder(style)
	} else {
		bc.Left.Style = style
	}
	return bc
}

// SetRight sets the right border style.
func (bc *Borders) SetRight(style BorderStyle) *Borders {
	if bc.Right == nil {
		bc.Right = NewBorder(style)
	} else {
		bc.Right.Style = style
	}
	return bc
}

// SetTop sets the top border style.
func (bc *Borders) SetTop(style BorderStyle) *Borders {
	if bc.Top == nil {
		bc.Top = NewBorder(style)
	} else {
		bc.Top.Style = style
	}
	return bc
}

// SetBottom sets the bottom border style.
func (bc *Borders) SetBottom(style BorderStyle) *Borders {
	if bc.Bottom == nil {
		bc.Bottom = NewBorder(style)
	} else {
		bc.Bottom.Style = style
	}
	return bc
}

// SetInner creates inner border configuration with the same style for all edges.
// Note: When applied to a table column, this will only be applied if the column has no sub-columns.
func (bc *Borders) SetInner(style BorderStyle) *Borders {
	border := NewBorder(style)
	bc.Inner = &Borders{
		Left:   border,
		Right:  border,
		Top:    border,
		Bottom: border,
	}
	return bc
}

// Style represents comprehensive styling configuration.
type Style struct {
	Bold            bool      // Whether text should be bold
	Italic          bool      // Whether text should be italic
	Underline       string    // Underline style (format-specific values)
	TextColor       string    // Text color (usually hex format: "#RRGGBB")
	BackgroundColor string    // Background color (usually hex format: "#RRGGBB")
	FontSize        float64   // Font size in points
	FontFamily      string    // Font family name (e.g., "Arial", "Times New Roman")
	Alignment       Alignment // Text alignment
}

// Alignment represents the alignment options for content.
type Alignment int

const (
	AlignmentNone         Alignment = iota // Use default alignment
	AlignmentLeft                          // Left-aligned text, top-aligned vertically
	AlignmentCenter                        // Center-aligned text, top-aligned vertically
	AlignmentRight                         // Right-aligned text, top-aligned vertically
	AlignmentTop                           // Left-aligned text, top-aligned vertically
	AlignmentMiddle                        // Left-aligned text, center-aligned vertically
	AlignmentBottom                        // Left-aligned text, bottom-aligned vertically
	AlignmentCenterMiddle                  // Center-aligned text, center-aligned vertically
	AlignmentLeftMiddle                    // Left-aligned text, center-aligned vertically
	AlignmentRightMiddle                   // Right-aligned text, center-aligned vertically
)

// GetAlignmentValues converts Alignment enum to horizontal and vertical alignment strings.
func (a Alignment) GetAlignmentValues() (horizontal, vertical string) {
	switch a {
	case AlignmentLeft:
		return "left", "top"
	case AlignmentCenter:
		return "center", "top"
	case AlignmentRight:
		return "right", "top"
	case AlignmentTop:
		return "left", "top"
	case AlignmentMiddle:
		return "left", "center"
	case AlignmentBottom:
		return "left", "bottom"
	case AlignmentCenterMiddle:
		return "center", "center"
	case AlignmentLeftMiddle:
		return "left", "center"
	case AlignmentRightMiddle:
		return "right", "center"
	default:
		return "left", "top" // Default alignment for unspecified cases
	}
}
