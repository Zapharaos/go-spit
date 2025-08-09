package table

import (
	"fmt"
	"strings"
)

// TODO : structs must be exported, functions must be kept private => builder pattern ?

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
		maxDepth := t.Columns.GetMaxDepth()
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

type Data map[string]interface{}

type DataSlice []Data

// Lookup looks up a key in a nested map structure.
func (d Data) Lookup(ks ...string) (rval interface{}, err error) {
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
		return d.Lookup(ks[1:]...)
	}
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

// HasSubColumns checks if this column contains nested sub-columns.
func (c Column) HasSubColumns() bool {
	return len(c.Columns) > 0
}

// GetColumnCount returns the total number of leaf columns represented by this column.
// For leaf columns, this returns 1. For parent columns, it recursively counts
// all leaf columns within the hierarchy.
func (c Column) GetColumnCount() int {
	if c.HasSubColumns() {
		// Recursively count all leaf columns in sub-columns
		total := 0
		for _, subCol := range c.Columns {
			total += subCol.GetColumnCount()
		}
		return total
	}
	return 1 // Leaf column represents exactly one data column
}

// Columns represents a collection of column definitions.
type Columns []Column

// GetTotalColumnCount calculates the total number of leaf columns in the collection.
func (c Columns) GetTotalColumnCount() int {
	total := 0
	for _, column := range c {
		total += column.GetColumnCount()
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

// GetFlattenedColumns returns a flattened list of all leaf columns.
// This traverses the entire column hierarchy and extracts only the columns
// that actually contain data (those without sub-columns).
func (c Columns) GetFlattenedColumns() []Column {
	var flattened []Column
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

// getParentColumnByIndex traverses the column hierarchy to find the parent column
// for a given column index.
func (c Columns) getParentColumnByIndex(colIndex int) *Column {
	// Use a helper function to maintain state during recursive traversal
	var helper func(cols Columns, targetIdx int, parent *Column, currentIdx *int) *Column
	helper = func(cols Columns, targetIdx int, parent *Column, currentIdx *int) *Column {
		for i := 0; i < len(cols); i++ {
			col := &cols[i]
			if col.HasSubColumns() {
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

			if col.HasSubColumns() {
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
