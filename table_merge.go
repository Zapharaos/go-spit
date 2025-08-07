package go_spit

import (
	"fmt"
)

// CellOperations defines the interface for cell operations that need to be passed to table merging functions.
// This interface abstracts the underlying cell manipulation operations, allowing the merging logic
// to work with different export formats (XLSX, CSV, etc.) by providing format-specific implementations.
type CellOperations interface {
	// GetCellValue retrieves the value from a cell at the specified coordinates.
	// Parameters:
	//   - col: 1-based column index
	//   - row: 1-based row index
	// Returns the cell value as a string and any error encountered.
	GetCellValue(col, row int) (string, error)

	// SetCellValue sets a value in a cell at the specified coordinates.
	// Parameters:
	//   - col: 1-based column index
	//   - row: 1-based row index
	//   - value: the value to set (can be any type)
	// Returns any error encountered during the operation.
	SetCellValue(col, row int, value interface{}) error

	// MergeCells merges cells from (startCol, startRow) to (endCol, endRow) inclusive.
	// The merged range will span from the top-left cell to the bottom-right cell.
	// Parameters:
	//   - startCol, startRow: top-left corner of the merge range (1-based)
	//   - endCol, endRow: bottom-right corner of the merge range (1-based)
	// Returns any error encountered during the merge operation.
	MergeCells(startCol, startRow, endCol, endRow int) error

	// IsCellMerged checks if a cell is part of any merged range.
	// Parameters:
	//   - col: 1-based column index
	//   - row: 1-based row index
	// Returns true if the cell is part of a merged range, false otherwise.
	IsCellMerged(col, row int) bool

	// IsCellMergedHorizontally checks if a cell is part of a horizontally merged range.
	// This is useful for determining if a cell spans multiple columns.
	// Parameters:
	//   - col: 1-based column index
	//   - row: 1-based row index
	// Returns true if the cell is part of a horizontal merge, false otherwise.
	IsCellMergedHorizontally(col, row int) bool

	// GetColumnLetter converts a 1-based column index to Excel column letter notation.
	// This is primarily used for debugging and logging purposes.
	// Parameters:
	//   - col: 1-based column index
	// Returns the column letter(s) (e.g., "A", "B", "AA", "AB").
	GetColumnLetter(col int) string

	// ProcessValue processes a raw value according to a format specification.
	// This handles type conversion, formatting, and any other value transformations.
	// Parameters:
	//   - value: the raw value to process
	//   - format: format specification string (e.g., date format, number format)
	// Returns the processed value and any error encountered.
	ProcessValue(value interface{}, format string) (interface{}, error)
}

// processTableCellMerging is the main entry point for table-based cell merging logic.
// It orchestrates both vertical (column-based) and horizontal (row-based) merging operations
// according to the table's configuration. This function processes the entire table systematically:
//
// 1. First, it handles vertical merging for each column individually
// 2. Then, it processes horizontal merging for each data row
// 3. Rows with specific merge configurations are handled with their custom settings
func (t *Table) processTableCellMerging(cellOps CellOperations) error {
	// Calculate where data rows start (after headers, if present)
	dataStartRow := t.getDataStartRow()

	// Phase 1: Process vertical merging for each flattened column
	// We use flattened columns because merging only applies to leaf columns
	// (columns that don't have sub-columns)
	for actualColIndex, column := range t.Columns.getFlattenedColumns() {
		// Column indices are 1-based, so we add 1 to the 0-based slice index
		if err := t.executeVerticalMergingForColumn(column, actualColIndex+1, dataStartRow, cellOps); err != nil {
			// Log the error but continue processing other columns
			L().Warn("Failed to execute vertical merging for column", Error(err))
		}
	}

	// Phase 2: Process horizontal merging for each data row
	for rowIndex, item := range t.Data {
		// Check if this row has custom merge configuration
		if rc, exists := t.RowConfigs[rowIndex]; exists && rc.Merge != nil && len(rc.Merge.Horizontal) > 0 {
			// Row has custom horizontal merge settings - process it with those settings
			if err := t.executeHorizontalMergingForRow(item, t.Columns, rowIndex, 1, &rc, cellOps); err != nil {
				return fmt.Errorf("failed to apply row horizontal merging: %w", err)
			}
			continue // Skip standard processing for this row
		}

		// Standard horizontal merging processing
		// Convert data row index to actual sheet row number
		rowNum := rowIndex + dataStartRow
		if err := t.executeHorizontalMergingForRow(item, t.Columns, rowNum, 1, nil, cellOps); err != nil {
			// Log the error but continue processing other rows
			L().Warn("Failed to execute horizontal merging for row", Int("row", rowNum), Error(err))
		}
	}

	return nil
}

// executeVerticalMergingForColumn handles vertical cell merging for a single column.
// This function identifies consecutive rows with values that meet the merge conditions
// and executes the merge operations.
func (t *Table) executeVerticalMergingForColumn(column Column, actualColIndex int, dataStartRow int, cellOps CellOperations) error {
	// Check if this column has vertical merge configuration
	mergeConfig := column.Merge
	if mergeConfig == nil || len(mergeConfig.Vertical) == 0 {
		return nil // No vertical merging configured for this column
	}

	// Step 1: Analyze the column data and identify merge ranges
	// This examines all data rows and finds consecutive groups that should be merged
	mergeRanges := t.findVerticalMergeRanges(actualColIndex, column.Name, column.Format, column.Merge.Vertical, cellOps)

	// Step 2: Execute merge operations for each identified range
	for _, mr := range mergeRanges {
		if len(mr) < 2 {
			continue // Skip single-cell ranges (nothing to merge)
		}

		// Convert data row indices to actual sheet row numbers
		startRow := mr[0] + dataStartRow
		endRow := mr[len(mr)-1] + dataStartRow

		// Execute the vertical merge operation
		if err := cellOps.MergeCells(actualColIndex, startRow, actualColIndex, endRow); err != nil {
			// Log detailed error information for debugging
			colLetter := cellOps.GetColumnLetter(actualColIndex)
			L().Warn("Failed to merge cells vertically",
				String("range", fmt.Sprintf("%s%d:%s%d", colLetter, startRow, colLetter, endRow)),
				Error(err))
			continue // Continue with other ranges even if one fails
		}
	}

	return nil
}

// findVerticalMergeRanges identifies ranges of consecutive rows that should be merged vertically.
func (t *Table) findVerticalMergeRanges(colIndex int, fieldName string, format string, conditions []MergeCondition, cellOps CellOperations) [][]int {
	var mergeRanges [][]int   // Collection of merge ranges to return
	var currentRange []int    // Current range being built
	var lastValue interface{} // Previous row's processed value for comparison

	// Iterate through each data row to analyze values and build ranges
	for rowIndex, item := range t.Data {
		// Skip rows that have custom vertical merge configurations
		// These are handled separately to avoid conflicts
		if rc, exists := t.RowConfigs[rowIndex]; exists && rc.Merge != nil {
			continue
		}

		// Check if this specific cell is marked as non-mergeable
		// This allows fine-grained control over which cells can be merged
		if cc, exists := t.CellConfigs[colIndex][rowIndex]; exists && !cc.Mergeable {
			continue
		}

		// Extract the raw value from the data item for this column
		value, err := item.Lookup(fieldName)
		if err != nil {
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
		processedValue, err := cellOps.ProcessValue(value, format)
		if err != nil {
			continue // Skip this row if value processing fails
		}

		if rowIndex == 0 {
			// First row - initialize the range tracking
			currentRange = []int{rowIndex}
			lastValue = processedValue
		} else {
			// Compare current value with previous value using merge conditions
			shouldMerge := evaluateMergeConditions(lastValue, processedValue, conditions)

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

// executeHorizontalMergingForRow processes horizontal cell merging for a single row.
// This function handles the complexity of nested column structures by using a two-pass approach:
//
// Pass 1: Recursively process all sub-columns from the deepest level up
// Pass 2: Process merging at the current column level
//
// This recursive approach ensures that sub-columns are merged before their parent columns,
// which is essential for proper hierarchical merging behavior. The function works with
// any depth of column nesting and respects the merge configurations at each level.
func (t *Table) executeHorizontalMergingForRow(item Data, columns Columns, rowNum int, startColIndex int, optRowConfig *RowConfig, cellOps CellOperations) error {
	currentColIndex := startColIndex

	// Pass 1: Recursively process all sub-columns (deepest level first)
	// This ensures that nested columns are fully processed before we handle
	// merging at the parent level
	for _, column := range columns {
		if column.hasSubColumns() {
			// Recursively process the sub-columns of this column
			if err := t.executeHorizontalMergingForRow(item, column.Columns, rowNum, currentColIndex, optRowConfig, cellOps); err != nil {
				return err
			}
			// Move to the next column group, accounting for all sub-columns
			currentColIndex += column.getColumnCount()
		} else {
			// Leaf column - just move to the next column
			currentColIndex++
		}
	}

	// Pass 2: Process horizontal merging at the current hierarchical level
	// Now that all sub-columns have been processed, we can safely evaluate
	// merging at this level
	return t.executeHorizontalMergingAtColumnLevel(item, columns, rowNum, startColIndex, optRowConfig, cellOps)
}

// executeHorizontalMergingAtColumnLevel handles horizontal merging at a specific column hierarchy level.
// This function examines consecutive columns to find groups that should be merged horizontally.
// The function respects cell-level merge restrictions.
func (t *Table) executeHorizontalMergingAtColumnLevel(item Data, columns Columns, rowNum int, startColIndex int, optRowConfig *RowConfig, cellOps CellOperations) error {
	currentColIndex := startColIndex

	// Process each column at this level
	for i := 0; i < len(columns); i++ {
		column := columns[i]

		// Check if the current cell is allowed to be merged
		dataRowIndex := t.getDataIndexFromRowIndex(rowNum)
		if cc, exists := t.CellConfigs[currentColIndex-1][dataRowIndex]; exists && !cc.Mergeable {
			currentColIndex++
			continue // Skip non-mergeable cells
		}

		// Skip columns without horizontal merge configuration
		if column.Merge == nil || len(column.Merge.Horizontal) == 0 {
			currentColIndex += column.getColumnCount()
			continue
		}

		// Determine the effective value for this column
		// This handles both leaf columns (direct data values) and parent columns (merged sub-column values)
		currentValue, err := t.findEffectiveColumnValue(item, column, rowNum, currentColIndex, cellOps)
		if err != nil {
			return fmt.Errorf("error finding effective value for column %s: %w", column.Name, err)
		}

		// Initialize merge group with the current column
		mergeGroup := []columnMergeInfo{{
			column:   column,
			startCol: currentColIndex,
			endCol:   currentColIndex + column.getColumnCount() - 1,
			value:    currentValue,
		}}

		// Look ahead to find consecutive columns that should merge with this one
		nextColIndex := currentColIndex + column.getColumnCount()
		for j := i + 1; j < len(columns); j++ {
			// Check if the next cell is mergeable
			if cc, exists := t.CellConfigs[nextColIndex-1][dataRowIndex]; exists && !cc.Mergeable {
				break // Stop if we encounter a non-mergeable cell
			}

			nextColumn := columns[j]
			nextValue, err := t.findEffectiveColumnValue(item, nextColumn, rowNum, nextColIndex, cellOps)
			if err != nil {
				return fmt.Errorf("error finding effective value for column %s: %w", nextColumn.Name, err)
			}

			// Determine which merge conditions to use
			merge := column.Merge
			if optRowConfig != nil && optRowConfig.Merge != nil {
				// Use row-specific merge conditions if available (they take precedence)
				merge = optRowConfig.Merge
			}

			// Check if the next column should merge with the current group
			shouldMerge := evaluateMergeConditions(currentValue, nextValue, merge.Horizontal)
			if shouldMerge {
				// Add the column to the merge group
				mergeGroup = append(mergeGroup, columnMergeInfo{
					column:   nextColumn,
					startCol: nextColIndex,
					endCol:   nextColIndex + nextColumn.getColumnCount() - 1,
					value:    nextValue,
				})
				nextColIndex += nextColumn.getColumnCount()
				i = j // Skip the merged column in the outer loop
			} else {
				break // Stop looking when we find a non-matching column
			}
		}

		// Execute merge operation if we have multiple columns to merge
		if len(mergeGroup) > 1 {
			if err = t.executeMergeCellsHorizontally(mergeGroup, rowNum, cellOps); err != nil {
				L().Warn("Failed to execute horizontal merge", Error(err))
			}
		}

		// Move to the next unprocessed column
		currentColIndex = nextColIndex
	}

	return nil
}

// findEffectiveColumnValue determines the effective value for a column in merge operations.
// This function handles the complexity of hierarchical columns by applying different logic
// based on whether the column has sub-columns:
//
// For leaf columns (no sub-columns):
//   - Extract the value directly from the data item
//   - Process the value according to the column's format specification
//
// For parent columns (with sub-columns):
//   - Check if all sub-columns have been merged into a single cell
//   - If merged, get the value from the merged cell
//   - If not merged, return empty to prevent merging at this level
//
// This approach ensures that merging respects the hierarchical structure of columns.
func (t *Table) findEffectiveColumnValue(item Data, column Column, rowNum int, startColIndex int, cellOps CellOperations) (string, error) {
	if !column.hasSubColumns() {
		// Leaf column: get the value directly from the data source
		value, _ := item.Lookup(column.Name)

		// Process the value according to the column's format specification
		// This ensures consistent formatting for merge comparison
		processedValue, err := cellOps.ProcessValue(value, column.Format)
		if err != nil {
			return "", fmt.Errorf("error processing value for column %s: %w", column.Name, err)
		}

		return fmt.Sprintf("%v", processedValue), nil
	}

	// Parent column: check if all sub-columns have been merged into one cell
	subColumnCount := column.getColumnCount()
	if subColumnCount == 1 {
		// All sub-columns were merged, so get the value from the merged cell
		// This represents the effective value for the entire column group
		value, err := cellOps.GetCellValue(startColIndex, rowNum)
		if err != nil {
			return "", err
		}
		return value, nil
	}

	// Sub-columns are not fully merged, so this parent column shouldn't merge
	// Returning empty string prevents merging at this level
	return "", nil
}

// executeMergeCellsHorizontally executes the horizontal merge operation for a group of columns.
// This function takes a group of consecutive columns that have been determined to be merge-compatible.
func (t *Table) executeMergeCellsHorizontally(mergeGroup []columnMergeInfo, rowNum int, cellOps CellOperations) error {
	if len(mergeGroup) < 2 {
		return nil // Nothing to merge with fewer than 2 columns
	}

	// Extract the range boundaries from the merge group
	startCol := mergeGroup[0].startCol             // First column in the group
	endCol := mergeGroup[len(mergeGroup)-1].endCol // Last column in the group

	// Delegate to the utility function to perform the actual merge
	return t.mergeCellRange(startCol, endCol, rowNum, cellOps)
}

// mergeCellRange is a utility function that executes a horizontal cell merge operation.
func (t *Table) mergeCellRange(startCol, endCol, rowNum int, cellOps CellOperations) error {
	if startCol >= endCol {
		return nil // Invalid range - nothing to merge
	}

	// Execute the merge operation using the cell operations interface
	if err := cellOps.MergeCells(startCol, rowNum, endCol, rowNum); err != nil {
		return fmt.Errorf("failed to merge cells from (%d,%d) to (%d,%d): %w", startCol, rowNum, endCol, rowNum, err)
	}

	return nil
}

// columnMergeInfo holds information about a column that's part of a merge group.
// This struct is used during horizontal merge processing to track the details
// of each column that should be merged together.
type columnMergeInfo struct {
	column   Column // The column definition containing merge rules and formatting
	startCol int    // Starting column index in the spreadsheet (1-based)
	endCol   int    // Ending column index in the spreadsheet (1-based)
	value    string // The effective value used for merge comparison
}
