// table_logic_merge.go - Table cell merging logic.
//
// This file implements algorithms and operations for merging cells in exported tables (vertical, horizontal, and header merging).
// Errors in merging are logged as warnings, allowing export to continue with best-effort merging.
// Used by spreadsheet backends to support advanced cell merging features in exported files.

package spit

import (
	"fmt"
)

// ProcessMerging applies all cell merging operations to the table.
// Handles header, vertical, and horizontal merging in order. Errors are logged and processing continues for best-effort merging.
func (t *Table) ProcessMerging(ops TableOperations) error {
	// Process header merging first
	if t.WriteHeader && len(t.Columns) > 0 {
		if err := t.executeHeaderMerging(ops); err != nil {
			return fmt.Errorf("failed to process header merging: %w", err)
		}
	}

	// Calculate where data rows start (after headers, if present)
	dataStartRow := t.GetDataStartRow()

	// Process vertical merging for each flattened column (leaf columns only)
	for actualColIndex, column := range t.Columns.GetFlattenedColumns() {
		// Column indices are 1-based, so we add 1 to the 0-based slice index
		if err := t.executeVerticalMerging(column, actualColIndex+1, dataStartRow, ops); err != nil {
			// Log the error but continue processing other columns
			L().Warn("Failed to process column for vertical merging", Error(err))
		}
	}

	// Process horizontal merging for each data row
	for rowIndex, item := range t.Data {
		// Convert data row index to actual sheet row number
		rowNum := rowIndex + dataStartRow

		// Retrieve any custom row options for this row
		rc, exists := t.RowOptionsMap[rowIndex]

		// If row has custom horizontal merge settings, process with those settings
		if exists && rc.Merge != nil && len(rc.Merge.Horizontal) > 0 {
			if err := t.executeHorizontalMerging(item, t.Columns, rowNum, 1, &rc, ops); err != nil {
				return fmt.Errorf("failed to apply row horizontal merging: %w", err)
			}
			continue // Skip standard processing for this row
		}

		// If row is marked as non-mergeable, skip merging
		if exists && !rc.Mergeable {
			continue
		}

		// Standard horizontal merging processing
		if err := t.executeHorizontalMerging(item, t.Columns, rowNum, 1, nil, ops); err != nil {
			L().Warn("Failed to execute horizontal merging for row", Int("row", rowNum), Error(err))
		}
	}

	return nil
}

// executeHeaderMerging applies merging operations to header cells.
// Handles hierarchical header merging for multi-level headers.
func (t *Table) executeHeaderMerging(ops TableOperations) error {
	maxDepth := t.Columns.GetMaxDepth()
	if maxDepth <= 1 {
		return nil // No merging needed for single-level headers
	}

	// Process hierarchical header merging recursively
	return t.processHeaderMergingRecursive(t.Columns, 1, maxDepth, 1, ops)
}

// processHeaderMergingRecursive processes header merging for hierarchical columns
func (t *Table) processHeaderMergingRecursive(columns Columns, currentRow, maxDepth, startCol int, ops TableOperations) error {
	currentCol := startCol

	for _, column := range columns {
		if column.HasSubColumns() {
			// Merge horizontally across sub-columns
			columnSpan := column.GetColumnCount()
			endCol := currentCol + columnSpan - 1
			if endCol > currentCol {
				if err := ops.MergeCells(currentCol, currentRow, endCol, currentRow); err != nil {
					L().Warn("Failed to merge header cells horizontally",
						Int("startCol", currentCol),
						Int("endCol", endCol),
						Int("row", currentRow),
						Error(err))
				}
			}

			// Recursively process sub-columns for next row level
			if currentRow < maxDepth {
				if err := t.processHeaderMergingRecursive(column.Columns, currentRow+1, maxDepth, currentCol, ops); err != nil {
					return err
				}
			}
			currentCol += columnSpan
		} else {
			// Merge vertically for leaf columns that span multiple header rows
			if currentRow < maxDepth {
				if err := ops.MergeCells(currentCol, currentRow, currentCol, maxDepth); err != nil {
					L().Warn("Failed to merge header cells vertically",
						Int("col", currentCol),
						Int("startRow", currentRow),
						Int("endRow", maxDepth),
						Error(err))
				}
			}
			currentCol++
		}
	}

	return nil
}

// executeVerticalMerging applies vertical cell merging for a single column.
// Only columns with vertical merge configuration are processed.
func (t *Table) executeVerticalMerging(column Column, actualColIndex int, dataStartRow int, ops TableOperations) error {
	if column.Merge == nil || len(column.Merge.Vertical) == 0 {
		return nil
	}

	// Analyze the column data and identify merge ranges
	mergeRanges := t.findVerticalMergeRanges(actualColIndex, column.Name, column.Format, column.Merge.Vertical, ops)

	// Execute merge operations for each identified range
	for _, mr := range mergeRanges {
		if len(mr) < 2 {
			continue // Skip single-cell ranges (nothing to merge)
		}

		// Convert data row indices to actual sheet row numbers
		startRow := mr[0] + dataStartRow
		endRow := mr[len(mr)-1] + dataStartRow

		// Execute the vertical merge operation
		if err := ops.MergeCells(actualColIndex, startRow, actualColIndex, endRow); err != nil {
			L().Warn("Failed to merge cells vertically",
				Int("col", actualColIndex),
				Int("startRow", startRow),
				Int("endRow", endRow),
				Error(err))
		}
	}

	return nil
}

// findVerticalMergeRanges identifies ranges of consecutive rows that should be merged vertically.
// Returns a slice of ranges (each range is a slice of row indices).
func (t *Table) findVerticalMergeRanges(colIndex int, fieldName string, format string, conditions MergeConditions, ops TableOperations) [][]int {
	var mergeRanges [][]int   // Collection of merge ranges to return
	var currentRange []int    // Current range being built
	var lastValue interface{} // Previous row's processed value for comparison

	// Iterate through each data row to analyze values and build ranges
	for rowIndex, item := range t.Data {
		// Skip rows that have disabled merging or have custom vertical merge configurations
		// Custom configurations are handled separately to avoid conflicts
		if rc, exists := t.RowOptionsMap[rowIndex]; exists && (!rc.Mergeable || rc.Merge != nil) {
			// Can't get value for this row - end current range if it exists
			if len(currentRange) > 1 {
				mergeRanges = append(mergeRanges, currentRange)
			}
			currentRange = nil
			lastValue = nil
			continue
		}

		// Check if this specific cell is marked as non-mergeable
		// This allows fine-grained control over which cells can be merged
		if cc, exists := t.CellOptionsMap[colIndex][rowIndex]; exists && !cc.Mergeable {
			// Can't get value for this row - end current range if it exists
			if len(currentRange) > 1 {
				mergeRanges = append(mergeRanges, currentRange)
			}
			currentRange = nil
			lastValue = nil
			continue
		}

		// Extract the raw value from the data item for this column
		value, err, found := item.Lookup(fieldName)
		if err != nil || !found {
			// Can't get value for this row - end current range if it exists
			if len(currentRange) > 1 {
				mergeRanges = append(mergeRanges, currentRange)
			}
			currentRange = nil
			lastValue = nil
			continue
		}

		// Process the value according to the column's format specification
		// This ensures consistent formatting for merge comparison
		processedValue, err := ops.ProcessValue(value, format)
		if err != nil {
			continue // Skip this row if value processing fails
		}

		if rowIndex == 0 {
			// First row - initialize the range tracking
			currentRange = []int{rowIndex}
			lastValue = processedValue
		} else {
			// Compare current value with previous value using merge conditions
			shouldMerge := conditions.ValuesShouldMerge(lastValue, processedValue)

			if shouldMerge {
				// Values should merge - add current row to the range
				if len(currentRange) == 0 {
					// Start a new range including the previous row
					currentRange = []int{rowIndex - 1, rowIndex}
				} else {
					// Extend the existing range
					currentRange = append(currentRange, rowIndex)
				}
			} else {
				// Values don't merge - finalize current range and start new one
				if len(currentRange) > 1 {
					// Save the completed range (only if it has multiple rows)
					mergeRanges = append(mergeRanges, currentRange)
				}
				// Start a new potential range with the current row
				currentRange = []int{rowIndex}
			}

			// Update the comparison value for the next iteration
			lastValue = processedValue
		}
	}

	// Don't forget to add the final range if it contains multiple rows
	if len(currentRange) > 1 {
		mergeRanges = append(mergeRanges, currentRange)
	}
	return mergeRanges
}

// executeHorizontalMerging processes horizontal cell merging for a single row.
// This function handles merging cells across columns within a row based on merge conditions.
func (t *Table) executeHorizontalMerging(item Data, columns Columns, rowNum int, startColIndex int, rowOptions *RowOptions, ops TableOperations) error {
	if len(columns) == 0 {
		return nil
	}

	// Check if row has custom horizontal merge configuration that overrides column settings
	// Row-level merge conditions take precedence over individual column configurations
	if rowOptions != nil && rowOptions.Merge != nil && len(rowOptions.Merge.Horizontal) > 0 {
		// Use row-level merge conditions for all columns in this row
		mergeRanges := t.findHorizontalMergeRanges(item, columns, rowOptions.Merge.Horizontal, ops)
		t.applyHorizontalMerges(mergeRanges, rowNum, startColIndex, ops)
		return nil
	}

	// Process columns with individual merge configurations
	// Group consecutive columns that have compatible merge conditions to optimize processing
	flatColumns := columns.GetFlattenedColumns()
	var currentGroup []Column             // Current group of columns being processed
	var currentGroupStartIndex int        // Starting index of the current group
	var currentConditions MergeConditions // Merge conditions for the current group

	// Iterate through columns and group them by compatible merge conditions
	for colIndex, column := range flatColumns {
		// Extract horizontal merge conditions for this column
		var columnConditions MergeConditions
		if column.Merge != nil && len(column.Merge.Horizontal) > 0 {
			columnConditions = column.Merge.Horizontal
		}

		// Group logic: start new group or add to existing group based on compatibility
		if len(currentGroup) == 0 {
			// Initialize the first group with this column
			currentGroup = []Column{column}
			currentGroupStartIndex = colIndex
			currentConditions = columnConditions
		} else if currentConditions.AnyMatch(columnConditions) {
			// Merge conditions are compatible - add column to current group
			currentGroup = append(currentGroup, column)
		} else {
			// Merge conditions are incompatible - process current group and start new one
			if len(currentGroup) > 1 && len(currentConditions) > 0 {
				// Process the completed group only if it has merge conditions and multiple columns
				groupColumns := Columns(currentGroup)
				mergeRanges := t.findHorizontalMergeRanges(item, groupColumns, currentConditions, ops)
				t.applyHorizontalMerges(mergeRanges, rowNum, startColIndex+currentGroupStartIndex, ops)
			}

			// Start a new group with the current column
			currentGroup = []Column{column}
			currentGroupStartIndex = colIndex
			currentConditions = columnConditions
		}
	}

	// Process the final group if it contains mergeable columns
	if len(currentGroup) > 1 && len(currentConditions) > 0 {
		groupColumns := Columns(currentGroup)
		mergeRanges := t.findHorizontalMergeRanges(item, groupColumns, currentConditions, ops)
		t.applyHorizontalMerges(mergeRanges, rowNum, startColIndex+currentGroupStartIndex, ops)
	}

	return nil
}

// applyHorizontalMerges executes horizontal merge operations for identified merge ranges.
func (t *Table) applyHorizontalMerges(mergeRanges [][]int, rowNum, baseColIndex int, ops TableOperations) {
	// Process each identified merge range
	for _, mergeRange := range mergeRanges {
		if len(mergeRange) < 2 {
			continue // Skip single-column ranges (nothing to merge)
		}

		// Convert relative column indices to absolute spreadsheet column positions
		startCol := mergeRange[0] + baseColIndex
		endCol := mergeRange[len(mergeRange)-1] + baseColIndex

		// Execute the horizontal merge operation across the column range
		if err := ops.MergeCells(startCol, rowNum, endCol, rowNum); err != nil {
			// Log detailed error information for debugging and continue processing
			L().Warn("Failed to merge cells horizontally",
				Int("row", rowNum),
				Int("startCol", startCol),
				Int("endCol", endCol),
				Error(err))
		}
	}
}

// findHorizontalMergeRanges identifies ranges of consecutive columns that should be merged horizontally.
// This function analyzes column values within a single row and determines which adjacent columns
// contain values that meet the specified merge conditions, building ranges of columns to merge.
func (t *Table) findHorizontalMergeRanges(item Data, columns Columns, conditions MergeConditions, ops TableOperations) [][]int {
	var mergeRanges [][]int   // Collection of merge ranges to return
	var currentRange []int    // Current range being built
	var lastValue interface{} // Previous column's processed value for comparison

	// Use flattened columns since merging only applies to leaf columns
	flatColumns := columns.GetFlattenedColumns()

	// Iterate through each column to analyze values and build merge ranges
	for colIndex, column := range flatColumns {
		// Check if this specific cell is marked as non-mergeable
		// This allows fine-grained control over which cells can participate in merging
		if cc, exists := t.CellOptionsMap[colIndex+1]; exists {
			if cellOptions, cellExists := cc[0]; cellExists && !cellOptions.Mergeable {
				// Cell is not mergeable - finalize current range and skip this column
				if len(currentRange) > 1 {
					mergeRanges = append(mergeRanges, currentRange)
				}
				currentRange = nil
				lastValue = nil
				continue
			}
		}

		// Extract the raw value from the data item for this column
		value, err, found := item.Lookup(column.Name)
		if err != nil || !found {
			// Can't get value for this column - end current range if it exists
			if len(currentRange) > 1 {
				mergeRanges = append(mergeRanges, currentRange)
			}
			currentRange = nil
			lastValue = nil
			continue
		}

		// Process the value according to the column's format specification
		// This ensures consistent formatting for merge comparison
		processedValue, err := ops.ProcessValue(value, column.Format)
		if err != nil {
			// Use raw value if processing fails
			processedValue = value
		}

		if colIndex == 0 {
			// First column - initialize the range tracking
			currentRange = []int{colIndex}
			lastValue = processedValue
		} else {
			// Compare current value with previous value using merge conditions
			shouldMerge := conditions.ValuesShouldMerge(lastValue, processedValue)

			if shouldMerge {
				// Values should merge - add current column to the range
				if len(currentRange) == 0 {
					// Start a new range including the previous column
					currentRange = []int{colIndex - 1, colIndex}
				} else {
					// Extend the existing range
					currentRange = append(currentRange, colIndex)
				}
			} else {
				// Values don't merge - finalize current range and start new one
				if len(currentRange) > 1 {
					// Save the completed range (only if it has multiple columns)
					mergeRanges = append(mergeRanges, currentRange)
				}
				// Start a new potential range with the current column
				currentRange = []int{colIndex}
			}

			// Update the comparison value for the next iteration
			lastValue = processedValue
		}
	}

	// Don't forget to add the final range if it contains multiple columns
	if len(currentRange) > 1 {
		mergeRanges = append(mergeRanges, currentRange)
	}

	return mergeRanges
}
