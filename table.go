package go_spit

import (
	"fmt"
	"strings"
)

type Columns []Column

// Column represents a generic column configuration for exports
type Column struct {
	Name    string
	Label   string
	Format  string
	Merge   *MergeConfig
	Border  BorderConfig
	Style   *StyleConfig
	Columns Columns
}

// HasSubColumns checks if this column has sub-columns
func (c Column) HasSubColumns() bool {
	return len(c.Columns) > 0
}

// GetColumnCount returns the number of actual columns this Column represents
func (c Column) GetColumnCount() int {
	if c.HasSubColumns() {
		total := 0
		for _, subCol := range c.Columns {
			total += subCol.GetColumnCount()
		}
		return total
	}
	return 1
}

// GetTotalColumnCount calculates the total number of columns considering sub-columns
func (c Columns) GetTotalColumnCount() int {
	total := 0
	for _, column := range c {
		total += column.GetColumnCount()
	}
	return total
}

// GetColumnsLabels returns the labels from a slice of columns
func (c Columns) GetColumnsLabels() []string {
	labels := make([]string, 0, len(c))
	for _, column := range c {
		labels = append(labels, column.Label)
	}
	return labels
}

// GetFlattenedColumns returns a flattened list of all leaf columns (columns without sub-columns)
func (c Columns) GetFlattenedColumns() []Column {
	var flattened []Column
	for _, column := range c {
		if column.HasSubColumns() {
			flattened = append(flattened, column.Columns.GetFlattenedColumns()...)
		} else {
			flattened = append(flattened, column)
		}
	}
	return flattened
}

// GetMaxDepth returns the maximum depth of the column hierarchy
func (c Columns) GetMaxDepth() int {
	maxDepth := 1
	for _, column := range c {
		if column.HasSubColumns() {
			depth := 1 + column.Columns.GetMaxDepth()
			if depth > maxDepth {
				maxDepth = depth
			}
		}
	}
	return maxDepth
}

// GetParentColumnByIndex traverses the column hierarchy and returns the parent column for the given colIndex
func (c Columns) GetParentColumnByIndex(colIndex int) *Column {
	var helper func(cols Columns, targetIdx int, parent *Column, currentIdx *int) *Column
	helper = func(cols Columns, targetIdx int, parent *Column, currentIdx *int) *Column {
		for i := 0; i < len(cols); i++ {
			col := &cols[i]
			if col.HasSubColumns() {
				result := helper(col.Columns, targetIdx, col, currentIdx)
				if result != nil {
					return result
				}
			} else {
				if *currentIdx == targetIdx {
					return parent
				}
				*currentIdx++
			}
		}
		return nil
	}
	idx := 0
	return helper(c, colIndex, nil, &idx)
}

type RowConfigs map[int]RowConfig // Maps row index to RowConfig

// RowConfig represents configuration for a specific row in the export
type RowConfig struct {
	RowIndex int           // The index of the row (1-based)
	Border   *BorderConfig // Border for the row
	Merge    *MergeConfig  // Merge config for the row
	Style    *StyleConfig  // Style configuration for the row
}

type CellConfigs map[int]CellConfigsByRow // Maps col index to CellConfigsByRow
type CellConfigsByRow map[int]CellConfig  // Maps row index to CellConfig

type CellConfig struct {
	RowIndex  int
	ColIndex  int
	Border    *BorderConfig
	Style     *StyleConfig
	Mergeable bool
}

// MergeCondition defines possible conditions for merging cells
// e.g. identical, empty, custom, etc.
type MergeCondition string

const (
	MergeConditionIdentical MergeCondition = "identical"
	MergeConditionEmpty     MergeCondition = "empty"
)

// MergeConfig holds merge conditions for a column
// If Conditions is empty, no merging is applied
// If not, merge is applied for listed conditions
// Example: MergeConfig{Conditions: []MergeCondition{MergeConditionIdentical, MergeConditionEmpty}}
type MergeConfig struct {
	Vertical   []MergeCondition `json:"vertical,omitempty"`   // Vertical merge conditions (between rows)
	Horizontal []MergeCondition `json:"horizontal,omitempty"` // Horizontal merge conditions (between columns)
}

// AreMergeConditionsCompatible checks if two sets of merge conditions are compatible
func AreMergeConditionsCompatible(conditions1, conditions2 []MergeCondition) bool {
	for _, cond1 := range conditions1 {
		for _, cond2 := range conditions2 {
			if cond1 == cond2 {
				return true
			}
		}
	}
	return false
}

// EvaluateMergeConditions determines if values should be merged based on merge conditions
// This unified function handles all merge scenarios: cell-to-cell, horizontal, and vertical
func EvaluateMergeConditions(value1, value2 interface{}, conditions []MergeCondition) bool {
	if len(conditions) == 0 {
		return false
	}

	// Convert values to strings for comparison
	val1Str := strings.TrimSpace(fmt.Sprintf("%v", value1))
	val2Str := strings.TrimSpace(fmt.Sprintf("%v", value2))

	// Check for empty values
	isEmpty1 := val1Str == "" || val1Str == "<nil>"
	isEmpty2 := val2Str == "" || val2Str == "<nil>"

	for _, condition := range conditions {
		switch condition {
		case MergeConditionIdentical:
			if val1Str == val2Str && !isEmpty1 && !isEmpty2 {
				return true
			}
		case MergeConditionEmpty:
			if isEmpty1 && isEmpty2 {
				return true
			}
		}
	}
	return false
}

// BorderStyle represents the style of a border
type BorderStyle int

const (
	BorderStyleNone   BorderStyle = 0
	BorderStyleThin   BorderStyle = 1
	BorderStyleMedium BorderStyle = 2
	BorderStyleDashed BorderStyle = 3
	BorderStyleDotted BorderStyle = 4
	BorderStyleThick  BorderStyle = 5
	BorderStyleDouble BorderStyle = 6
)

// BorderSide represents configuration for one side of a border
type BorderSide struct {
	Style BorderStyle
}

// BorderConfig represents border configuration for a column
type BorderConfig struct {
	Left   *BorderSide
	Right  *BorderSide
	Top    *BorderSide
	Bottom *BorderSide
	Inner  *BorderConfig
}

// HasBorders checks if any borders are configured
func (bc BorderConfig) HasBorders() bool {
	return (bc.Left != nil && bc.Left.Style != BorderStyleNone) ||
		(bc.Right != nil && bc.Right.Style != BorderStyleNone) ||
		(bc.Top != nil && bc.Top.Style != BorderStyleNone) ||
		(bc.Bottom != nil && bc.Bottom.Style != BorderStyleNone)
}

// SetInner sets the inner border configuration to the same style for all sides
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

// NewBorderConfig creates a BorderConfig with the same style applied to all sides
func NewBorderConfig(style BorderStyle) BorderConfig {
	side := &BorderSide{Style: style}
	return BorderConfig{
		Left:   side,
		Right:  side,
		Top:    side,
		Bottom: side,
	}
}

type StyleConfig struct {
	Bold            bool
	Italic          bool
	Underline       string
	TextColor       string
	BackgroundColor string
	FontSize        float64
	FontFamily      string
	Alignment       CellAlignment
}

// CellAlignment represents the alignment options for cell content
type CellAlignment int

const (
	AlignmentNone CellAlignment = iota
	AlignmentLeft
	AlignmentCenter
	AlignmentRight
	AlignmentTop
	AlignmentMiddle
	AlignmentBottom
	AlignmentCenterMiddle
	AlignmentLeftMiddle
	AlignmentRightMiddle
)
