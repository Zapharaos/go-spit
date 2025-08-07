package go_spit

import (
	"fmt"
	"strings"
)

// Table represents a structured data table with configuration for export operations.
// It contains the actual data, column definitions, and various configuration options
// that control how the table should be rendered in different export formats.
type Table struct {
	Data          DataSlice   // The actual data rows to be exported
	Columns       Columns     // Column definitions including hierarchy and formatting
	RowConfigs    RowConfigs  // Optional row-specific configurations (styling, merging, borders)
	CellConfigs   CellConfigs // Optional cell-specific configurations for fine-grained control
	WriteHeader   bool        // Whether to generate headers from column definitions
	Limit         int64       // Maximum number of data rows to export (0 = no limit)
	ListSeparator string      // Separator used when rendering slice/array values as strings
}

// getDataStartRow calculates the starting row number for data based on header configuration.
// The calculation accounts for multi-level headers by determining the maximum depth
// of the column hierarchy and reserving that many rows for header content.
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

// getDataIndexFromRowIndex converts a row index to the corresponding data slice index.
// This conversion is necessary because the Data slice is 0-based, but export formats
// typically use 1-based row numbering.
func (t *Table) getDataIndexFromRowIndex(rowIndex int) int {
	if t.WriteHeader {
		// Adjust row index to account for header rows
		headerShift := t.getDataStartRow()
		return rowIndex - headerShift
	}
	return rowIndex
}

// Column represents a single column definition for table exports.
// Columns can be nested to create hierarchical structures, allowing for
// complex header layouts and grouped data organization.
type Column struct {
	Name    string       // Field name in the data source (for leaf columns)
	Label   string       // Display label for headers
	Format  string       // Format specification for value processing (e.g., date format)
	Merge   *MergeConfig // Optional merge configuration for this column
	Border  BorderConfig // Border styling configuration
	Style   *StyleConfig // Text and cell styling configuration
	Columns Columns      // Sub-columns for hierarchical structures
}

// hasSubColumns checks if this column contains nested sub-columns.
func (c Column) hasSubColumns() bool {
	return len(c.Columns) > 0
}

// getColumnCount returns the total number of leaf columns represented by this column.
// For leaf columns, this returns 1. For parent columns, it recursively counts
// all leaf columns within the hierarchy.
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

// Columns represents a collection of column definitions.
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

// getFlattenedColumns returns a flattened list of all leaf columns.
// This traverses the entire column hierarchy and extracts only the columns
// that actually contain data (those without sub-columns).
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

// getParentColumnByIndex traverses the column hierarchy to find the parent column
// for a given column index.
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

// getColumnIndex returns the index of a specific column within the hierarchy.
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

// RowConfigs maps row indices to their specific configurations.
type RowConfigs map[int]RowConfig

// RowConfig represents configuration options for a specific row in the export.
// This allows fine-grained control over individual rows, overriding default
// column-based settings when needed.
type RowConfig struct {
	RowIndex int           // The 0-based index of the row this configuration applies to
	Border   *BorderConfig // Optional border configuration for the entire row
	Merge    *MergeConfig  // Optional merge configuration that overrides column settings
	Style    *StyleConfig  // Optional style configuration for the entire row
}

// CellConfigs provides cell-level configuration mapping.
// The outer map keys are column indices, inner map keys are row indices.
type CellConfigs map[int]CellConfigsByRow

// CellConfigsByRow maps row indices to cell configurations for a specific column.
type CellConfigsByRow map[int]CellConfig

// CellConfig represents configuration options for a specific cell.
// This provides the finest level of control, allowing individual cells
// to override both column and row settings.
type CellConfig struct {
	RowIndex  int           // The 0-based row index of this cell
	ColIndex  int           // The 0-based column index of this cell
	Border    *BorderConfig // Optional border configuration for this cell
	Style     *StyleConfig  // Optional style configuration for this cell
	Mergeable bool          // Whether this cell can participate in merge operations
}

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

// MergeConfig holds merge conditions for columns and rows.
// It defines when and how cells should be merged based on their content.
// Empty conditions arrays mean no merging will be applied.
type MergeConfig struct {
	Vertical   []MergeCondition `json:"vertical,omitempty"`   // Conditions for merging cells vertically (between rows)
	Horizontal []MergeCondition `json:"horizontal,omitempty"` // Conditions for merging cells horizontally (between columns)
}

// areMergeConditionsCompatible checks if two sets of merge conditions share at least one common condition.
// This is used to determine if two cells or ranges can be merged together based on their configurations.
func areMergeConditionsCompatible(conditions1, conditions2 []MergeCondition) bool {
	for _, cond1 := range conditions1 {
		for _, cond2 := range conditions2 {
			if cond1 == cond2 {
				return true // Found a matching condition
			}
		}
	}
	return false // No compatible conditions found
}

// evaluateMergeConditions determines if two values should be merged based on the specified conditions.
func evaluateMergeConditions(value1, value2 interface{}, conditions []MergeCondition) bool {
	if len(conditions) == 0 {
		return false // No conditions specified - don't merge
	}

	// Convert values to strings for consistent comparison
	val1Str := strings.TrimSpace(fmt.Sprintf("%v", value1))
	val2Str := strings.TrimSpace(fmt.Sprintf("%v", value2))

	// Determine if values are considered empty
	isEmpty1 := val1Str == "" || val1Str == "<nil>"
	isEmpty2 := val2Str == "" || val2Str == "<nil>"

	// Evaluate each condition to see if any match
	for _, condition := range conditions {
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

// BorderStyle represents the visual style of cell borders.
// These constants correspond to common border styles available in spreadsheet applications.
type BorderStyle int

const (
	BorderStyleNone   BorderStyle = 0 // No border
	BorderStyleThin   BorderStyle = 1 // Thin solid line
	BorderStyleMedium BorderStyle = 2 // Medium thickness solid line
	BorderStyleDashed BorderStyle = 3 // Dashed line
	BorderStyleDotted BorderStyle = 4 // Dotted line
	BorderStyleThick  BorderStyle = 5 // Thick solid line
	BorderStyleDouble BorderStyle = 6 // Double line
)

// BorderSide represents the configuration for one side of a cell border.
type BorderSide struct {
	Style BorderStyle // The visual style to apply to this border side
}

// BorderConfig represents complete border configuration for a cell, row, or column.
type BorderConfig struct {
	Left   *BorderSide   // Left border configuration
	Right  *BorderSide   // Right border configuration
	Top    *BorderSide   // Top border configuration
	Bottom *BorderSide   // Bottom border configuration
	Inner  *BorderConfig // Inner borders for ranges (used in some contexts)
}

// hasBorders checks if any borders are configured in this BorderConfig.
func (bc BorderConfig) hasBorders() bool {
	return (bc.Left != nil && bc.Left.Style != BorderStyleNone) ||
		(bc.Right != nil && bc.Right.Style != BorderStyleNone) ||
		(bc.Top != nil && bc.Top.Style != BorderStyleNone) ||
		(bc.Bottom != nil && bc.Bottom.Style != BorderStyleNone)
}

// SetInner creates inner border configuration with the same style for all sides.
func (bc BorderConfig) SetInner(style BorderStyle) BorderConfig {
	side := &BorderSide{Style: style}
	bc.Inner = &BorderConfig{
		Left:   side,
		Right:  side,
		Top:    side,
		Bottom: side,
	}
	return bc
}

// NewBorderConfig creates a BorderConfig with the same style applied to all sides.
func NewBorderConfig(style BorderStyle) BorderConfig {
	side := &BorderSide{Style: style}
	return BorderConfig{
		Left:   side,
		Right:  side,
		Top:    side,
		Bottom: side,
	}
}

// StyleConfig represents comprehensive styling configuration for cells.
type StyleConfig struct {
	Bold            bool          // Whether text should be bold
	Italic          bool          // Whether text should be italic
	Underline       string        // Underline style (format-specific values)
	TextColor       string        // Text color (usually hex format: "#RRGGBB")
	BackgroundColor string        // Cell background color (usually hex format: "#RRGGBB")
	FontSize        float64       // Font size in points
	FontFamily      string        // Font family name (e.g., "Arial", "Times New Roman")
	Alignment       CellAlignment // Text alignment within the cell
}

// CellAlignment represents the alignment options for cell content.
type CellAlignment int

const (
	AlignmentNone         CellAlignment = iota // Use default alignment
	AlignmentLeft                              // Left-aligned text, top-aligned vertically
	AlignmentCenter                            // Center-aligned text, top-aligned vertically
	AlignmentRight                             // Right-aligned text, top-aligned vertically
	AlignmentTop                               // Left-aligned text, top-aligned vertically
	AlignmentMiddle                            // Left-aligned text, center-aligned vertically
	AlignmentBottom                            // Left-aligned text, bottom-aligned vertically
	AlignmentCenterMiddle                      // Center-aligned text, center-aligned vertically
	AlignmentLeftMiddle                        // Left-aligned text, center-aligned vertically
	AlignmentRightMiddle                       // Right-aligned text, center-aligned vertically
)

// getAlignmentValues converts CellAlignment enum to horizontal and vertical alignment strings.
func (ca CellAlignment) getAlignmentValues() (horizontal, vertical string) {
	switch ca {
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
