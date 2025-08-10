// table.go - Table abstraction and operations.
//
// This file defines the Table structure and the tableOperations interface, which provide a unified way to
// represent, manipulate, and export tabular data with hierarchical columns, formatting, and cell/row configuration.
// It allows for flexible data representation and export across different libraries.

package go_spit

import (
	"fmt"
	"strings"
)

// tableOperations defines table-specific operations interface.
// Implement this interface to support table manipulation for various libraries.
type tableOperations interface {
	// getTable returns the underlying Table struct for direct access/manipulation.
	getTable() *Table

	// getCellValue returns the value of a cell at the given column and row (1-based indices).
	// The implementation should convert indices to the appropriate cell reference and retrieve the value.
	getCellValue(col, row int) (string, error)

	// setCellValue sets the value of a cell at the given column and row.
	// The implementation should convert indices to the appropriate cell reference and set the value.
	setCellValue(col, row int, value interface{}) error

	// mergeCells merges a rectangular range of cells defined by start and end coordinates.
	mergeCells(startCol, startRow, endCol, endRow int) error

	// isCellMerged checks if a cell at the given column and row is part of any merged range.
	isCellMerged(col, row int) bool

	// isCellMergedHorizontally checks if a cell at the given column and row is merged horizontally (across columns).
	isCellMergedHorizontally(col, row int) bool

	// applyBorderToCell applies a border to a cell on the specified side ("left", "right", "top", "bottom").
	// The border style is defined by the Border parameter.
	applyBorderToCell(col, row int, side string, border *Border) error

	// applyBordersToRange applies borders to all cells in a rectangular range.
	// Each side of the range can have a different border style, as specified in the Borders parameter.
	applyBordersToRange(startCol, startRow, endCol, endRow int, borders Borders) error

	// hasExistingBorder checks if a cell at the given column and row has a border on the specified side.
	hasExistingBorder(col, row int, side string) bool

	// applyStyleToCell applies a style to a specific cell at the given column and row.
	applyStyleToCell(col, row int, style Style) error

	// applyStyleToRange Applies a style to a rectangular range of cells.
	applyStyleToRange(startCol, startRow, endCol, endRow int, style Style) error

	// getColumnLetter Returns the Excel-style column letter (e.g., "A", "B") for a given column index.
	getColumnLetter(col int) string

	// processValue Processes a value for output, applying formatting if needed.
	processValue(value interface{}, format string) (interface{}, error)
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
	ListSeparator  string         // Separator used when rendering slice/array values as strings
}

// getDataStartRow calculates the starting row number for data based on header configuration.
// Accounts for multi-level headers by reserving rows for each level of the column hierarchy.
func (t *Table) getDataStartRow() int {
	dataStartRow := 1
	if t.WriteHeader && len(t.Columns) > 0 {
		// Calculate the maximum depth of the column hierarchy
		// Each level of nesting requires its own header row
		maxDepth := t.Columns.getMaxDepth()
		dataStartRow = maxDepth + 1
	}
	return dataStartRow
}

// getDataIndexFromRowIndex converts a 1-based row index to the corresponding 0-based data slice index.
// Adjusts for header rows if present.
func (t *Table) getDataIndexFromRowIndex(rowIndex int) int {
	if t.WriteHeader {
		// Adjust row index to account for header rows
		headerShift := t.getDataStartRow()
		return rowIndex - headerShift
	}
	return rowIndex
}

// Data represents a single row of table data as a map from column name to value.
type Data map[string]interface{}

// DataSlice is a slice of Data rows.
type DataSlice []Data

// lookup recursively looks up a key in a nested map structure.
// Supports multi-level key access for hierarchical data.
func (d Data) lookup(ks ...string) (rval interface{}, err error) {
	var ok bool
	if len(ks) == 0 {
		return nil, fmt.Errorf("nestedMapLookup needs at least one key")
	}
	if rval, ok = d[strings.TrimSpace(ks[0])]; !ok {
		return nil, fmt.Errorf("key not found; remaining keys: %v", ks)
	} else if len(ks) == 1 {
		return rval, nil
	} else if d, ok = rval.(Data); !ok {
		return nil, fmt.Errorf("malformed structure at %#v", rval)
	} else {
		return d.lookup(ks[1:]...)
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
	Borders Borders     // Borders configuration
	Style   *Style      // Optional content style
	Columns Columns     // Sub-columns for hierarchical structures
}

// hasSubColumns returns true if this column contains nested sub-columns.
func (c Column) hasSubColumns() bool {
	return len(c.Columns) > 0
}

// getColumnCount returns the total number of leaf columns represented by this column.
// For leaf columns, returns 1. For parent columns, recursively counts all leaf columns in the hierarchy.
func (c Column) getColumnCount() int {
	if c.hasSubColumns() {
		// Recursively count all leaf columns in sub-columns
		total := 0
		for _, subCol := range c.Columns {
			total += subCol.getColumnCount()
		}
		return total
	}
	return 1 // Leaf column represents exactly one data column
}

// Columns represents a collection of column definitions, possibly hierarchical.
type Columns []Column

// getTotalColumnCount calculates the total number of leaf columns in the collection.
func (c Columns) getTotalColumnCount() int {
	total := 0
	for _, column := range c {
		total += column.getColumnCount()
	}
	return total
}

// getColumnsLabels returns the display labels from all columns in the collection.
func (c Columns) getColumnsLabels() []string {
	labels := make([]string, 0, len(c))
	for _, column := range c {
		labels = append(labels, column.Label)
	}
	return labels
}

// getFlattenedColumns returns a flattened list of all leaf columns in the hierarchy.
// Traverses the entire column structure and extracts only columns without sub-columns.
func (c Columns) getFlattenedColumns() []Column {
	var flattened []Column
	for _, column := range c {
		if column.hasSubColumns() {
			// Recursively flatten sub-columns
			flattened = append(flattened, column.Columns.getFlattenedColumns()...)
		} else {
			// Add leaf column directly
			flattened = append(flattened, column)
		}
	}
	return flattened
}

// getMaxDepth returns the maximum depth of the column hierarchy.
func (c Columns) getMaxDepth() int {
	maxDepth := 1
	for _, column := range c {
		if column.hasSubColumns() {
			// Recursively calculate depth including this level
			depth := 1 + column.Columns.getMaxDepth()
			if depth > maxDepth {
				maxDepth = depth
			}
		}
	}
	return maxDepth
}

// getParentColumnByIndex traverses the column hierarchy to find the parent column for a given leaf column index.
func (c Columns) getParentColumnByIndex(colIndex int) *Column {
	// Use a helper function to maintain state during recursive traversal
	var helper func(cols Columns, targetIdx int, parent *Column, currentIdx *int) *Column
	helper = func(cols Columns, targetIdx int, parent *Column, currentIdx *int) *Column {
		for i := 0; i < len(cols); i++ {
			col := &cols[i]
			if col.hasSubColumns() {
				// Recursively search in sub-columns with this column as parent
				result := helper(col.Columns, targetIdx, col, currentIdx)
				if result != nil {
					return result // Found the target in sub-columns
				}
			} else {
				// Check if this leaf column matches the target index
				if *currentIdx == targetIdx {
					return parent // Return the parent of the matching column
				}
				*currentIdx++ // Move to next leaf column
			}
		}
		return nil // Target not found in this branch
	}

	idx := 0
	return helper(c, colIndex, nil, &idx)
}

// getColumnIndex returns the index of a specific column within the hierarchy (1-based).
func (c Columns) getColumnIndex(target *Column) int {
	currentIndex := 1

	// Recursive helper function to traverse the hierarchy
	var findIndex func(cols Columns, target *Column, currentIdx *int) bool
	findIndex = func(cols Columns, target *Column, currentIdx *int) bool {
		for i := range cols {
			col := &cols[i]
			if col == target {
				return true // Found the target column
			}

			if col.hasSubColumns() {
				// Search recursively in sub-columns
				if findIndex(col.Columns, target, currentIdx) {
					return true
				}
			} else {
				// Move to next position for leaf columns
				*currentIdx++
			}
		}
		return false // Target not found in this branch
	}

	if findIndex(c, target, &currentIndex) {
		return currentIndex
	}

	return 0 // Column not found
}

// RowOptionsMap maps row indices to their specific options.
type RowOptionsMap map[int]RowOptions

// RowOptions represents option settings for a specific row in the export.
// This allows fine-grained control over individual rows, overriding default
// column-based settings when needed.
type RowOptions struct {
	RowIndex int         // The 0-based index of the row this option applies to
	Border   *Borders    // Optional border configuration for the entire row
	Merge    *MergeRules // Optional merge configuration that overrides column settings
	Style    *Style      // Optional style configuration for the entire row
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

// MergeRules holds merge conditions for columns and rows.
// It defines when and how cells should be merged based on their content.
// Empty conditions arrays mean no merging will be applied.
type MergeRules struct {
	Vertical   MergeConditions `json:"vertical,omitempty"`   // Conditions for merging cells vertically (between rows)
	Horizontal MergeConditions `json:"horizontal,omitempty"` // Conditions for merging cells horizontally (between columns)
}

// anyMatch checks if two sets of merge conditions share at least one common condition.
// This is used to determine if two cells or ranges can be merged together based on their configurations.
func (mc MergeConditions) anyMatch(mc2 []MergeCondition) bool {
	for _, cond1 := range mc {
		for _, cond2 := range mc2 {
			if cond1 == cond2 {
				return true // Found a matching condition
			}
		}
	}
	return false // No compatible conditions found
}

// valuesShouldMerge determines if two values should be merged based on the specified conditions.
func (mc MergeConditions) valuesShouldMerge(value1, value2 interface{}) bool {
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

// Borders represents all borders configuration for an entity.
type Borders struct {
	Left   *Border  // Left border configuration
	Right  *Border  // Right border configuration
	Top    *Border  // Top border configuration
	Bottom *Border  // Bottom border configuration
	Inner  *Borders // Inner borders for ranges (used in some contexts)
}

// hasBorders checks if any borders are configured in these Borders.
func (bc Borders) hasBorders() bool {
	return (bc.Left != nil && bc.Left.Style != BorderStyleNone) ||
		(bc.Right != nil && bc.Right.Style != BorderStyleNone) ||
		(bc.Top != nil && bc.Top.Style != BorderStyleNone) ||
		(bc.Bottom != nil && bc.Bottom.Style != BorderStyleNone)
}

// SetInner creates inner border configuration with the same style for all edges.
func (bc Borders) SetInner(style BorderStyle) Borders {
	border := &Border{Style: style}
	bc.Inner = &Borders{
		Left:   border,
		Right:  border,
		Top:    border,
		Bottom: border,
	}
	return bc
}

// NewBorderOptions creates a Borders with the same style applied to all edges.
func NewBorderOptions(style BorderStyle) Borders {
	border := &Border{Style: style}
	return Borders{
		Left:   border,
		Right:  border,
		Top:    border,
		Bottom: border,
	}
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

// getAlignmentValues converts Alignment enum to horizontal and vertical alignment strings.
func (a Alignment) getAlignmentValues() (horizontal, vertical string) {
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
