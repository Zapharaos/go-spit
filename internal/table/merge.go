package table

import (
	"fmt"
	"strings"

	"github.com/Zapharaos/go-spit/internal/logger"
)

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

// ProcessMerging applies cell merging operations to the table
func (t *Table) ProcessMerging(ops Operations) error {
	// Phase 1: Process header merging first
	if t.WriteHeader && len(t.Columns) > 0 {
		if err := t.executeHeaderMerging(ops); err != nil {
			return fmt.Errorf("failed to process header merging: %w", err)
		}
	}

	// Calculate where data rows start (after headers, if present)
	dataStartRow := t.getDataStartRow()

	// Process vertical merging for each flattened column
	// We use flattened columns because merging only applies to leaf columns
	for actualColIndex, column := range t.Columns.GetFlattenedColumns() {
		// Column indices are 1-based, so we add 1 to the 0-based slice index
		if err := t.executeVerticalMerging(column, actualColIndex+1, dataStartRow, ops); err != nil {
			// Log the error but continue processing other columns
			logger.L().Warn("Failed to process column for vertical merging", logger.Error(err))
		}
	}

	// Process horizontal merging for each data row
	for rowIndex, item := range t.Data {
		// Check if this row has custom merge configuration
		if rc, exists := t.RowConfigs[rowIndex]; exists && rc.Merge != nil && len(rc.Merge.Horizontal) > 0 {
			// Row has custom horizontal merge settings - process it with those settings
			if err := t.executeHorizontalMerging(item, t.Columns, rowIndex, 1, &rc, ops); err != nil {
				return fmt.Errorf("failed to apply row horizontal merging: %w", err)
			}
			continue // Skip standard processing for this row
		}

		// Standard horizontal merging processing
		// Convert data row index to actual sheet row number
		rowNum := rowIndex + dataStartRow
		if err := t.executeHorizontalMerging(item, t.Columns, rowNum, 1, nil, ops); err != nil {
			// Log the error but continue processing other rows
			logger.L().Warn("Failed to execute horizontal merging for row", logger.Int("row", rowNum), logger.Error(err))
		}
	}

	return nil
}

// executeHeaderMerging applies merging operations to header cells
func (t *Table) executeHeaderMerging(ops Operations) error {
	maxDepth := t.Columns.GetMaxDepth()
	if maxDepth <= 1 {
		return nil // No merging needed for single-level headers
	}

	// Process hierarchical header merging
	return t.processHeaderMergingRecursive(t.Columns, 1, maxDepth, 1, ops)
}

// processHeaderMergingRecursive processes header merging for hierarchical columns
func (t *Table) processHeaderMergingRecursive(columns Columns, currentRow, maxDepth, startCol int, ops Operations) error {
	currentCol := startCol

	for _, column := range columns {
		if column.HasSubColumns() {
			// Merge horizontally across sub-columns
			columnSpan := column.GetColumnCount()
			endCol := currentCol + columnSpan - 1
			if endCol > currentCol {
				if err := ops.MergeCells(currentCol, currentRow, endCol, currentRow); err != nil {
					logger.L().Warn("Failed to merge header cells horizontally",
						logger.Int("startCol", currentCol),
						logger.Int("endCol", endCol),
						logger.Int("row", currentRow),
						logger.Error(err))
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
					logger.L().Warn("Failed to merge header cells vertically",
						logger.Int("col", currentCol),
						logger.Int("startRow", currentRow),
						logger.Int("endRow", maxDepth),
						logger.Error(err))
				}
			}
			currentCol++
		}
	}

	return nil
}

// executeVerticalMergingForColumn handles vertical cell merging for a single column.
func (t *Table) executeVerticalMerging(column Column, actualColIndex int, dataStartRow int, ops Operations) error {
	// Check if this column has vertical merge configuration
	if column.Merge == nil || len(column.Merge.Vertical) == 0 {
		return nil
	}

	// Analyze the column data and identify merge ranges
	mergeRanges := t.findVerticalMergeRanges(actualColIndex, column.Name, column.Format, column.Merge.Vertical, ops)

	//  Execute merge operations for each identified range
	for _, mr := range mergeRanges {
		if len(mr) < 2 {
			continue // Skip single-cell ranges (nothing to merge)
		}

		// Convert data row indices to actual sheet row numbers
		startRow := mr[0] + dataStartRow
		endRow := mr[len(mr)-1] + dataStartRow

		// Execute the vertical merge operation
		if err := ops.MergeCells(actualColIndex, startRow, actualColIndex, endRow); err != nil {
			// Log detailed error information for debugging and continue processing
			logger.L().Warn("Failed to merge cells vertically",
				logger.Int("col", actualColIndex),
				logger.Int("startRow", startRow),
				logger.Int("endRow", endRow),
				logger.Error(err))
		}
	}

	return nil
}

// findVerticalMergeRanges identifies ranges of consecutive rows that should be merged vertically.
func (t *Table) findVerticalMergeRanges(colIndex int, fieldName string, format string, conditions []MergeCondition, ops Operations) [][]int {
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

// executeHorizontalMerging processes horizontal cell merging for a single row.
// This function handles merging cells across columns within a row based on merge conditions.
func (t *Table) executeHorizontalMerging(item Data, columns Columns, rowNum int, startColIndex int, optRowConfig *RowConfig, ops Operations) error {
	if len(columns) == 0 {
		return nil
	}

	// Check if row has custom horizontal merge configuration that overrides column settings
	// Row-level merge conditions take precedence over individual column configurations
	if optRowConfig != nil && optRowConfig.Merge != nil && len(optRowConfig.Merge.Horizontal) > 0 {
		// Use row-level merge conditions for all columns in this row
		mergeRanges := t.findHorizontalMergeRanges(item, columns, optRowConfig.Merge.Horizontal, ops)
		t.applyHorizontalMerges(mergeRanges, rowNum, startColIndex, ops)
		return nil
	}

	// Process columns with individual merge configurations
	// Group consecutive columns that have compatible merge conditions to optimize processing
	flatColumns := columns.GetFlattenedColumns()
	var currentGroup []Column              // Current group of columns being processed
	var currentGroupStartIndex int         // Starting index of the current group
	var currentConditions []MergeCondition // Merge conditions for the current group

	// Iterate through columns and group them by compatible merge conditions
	for colIndex, column := range flatColumns {
		// Extract horizontal merge conditions for this column
		var columnConditions []MergeCondition
		if column.Merge != nil && len(column.Merge.Horizontal) > 0 {
			columnConditions = column.Merge.Horizontal
		}

		// Group logic: start new group or add to existing group based on compatibility
		if len(currentGroup) == 0 {
			// Initialize the first group with this column
			currentGroup = []Column{column}
			currentGroupStartIndex = colIndex
			currentConditions = columnConditions
		} else if areMergeConditionsCompatible(currentConditions, columnConditions) {
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
func (t *Table) applyHorizontalMerges(mergeRanges [][]int, rowNum, baseColIndex int, ops Operations) {
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
			logger.L().Warn("Failed to merge cells horizontally",
				logger.Int("row", rowNum),
				logger.Int("startCol", startCol),
				logger.Int("endCol", endCol),
				logger.Error(err))
		}
	}
}

// findHorizontalMergeRanges identifies ranges of consecutive columns that should be merged horizontally.
// This function analyzes column values within a single row and determines which adjacent columns
// contain values that meet the specified merge conditions, building ranges of columns to merge.
func (t *Table) findHorizontalMergeRanges(item Data, columns Columns, conditions []MergeCondition, ops Operations) [][]int {
	var mergeRanges [][]int   // Collection of merge ranges to return
	var currentRange []int    // Current range being built
	var lastValue interface{} // Previous column's processed value for comparison

	// Use flattened columns since merging only applies to leaf columns
	flatColumns := columns.GetFlattenedColumns()

	// Iterate through each column to analyze values and build merge ranges
	for colIndex, column := range flatColumns {
		// Check if this specific cell is marked as non-mergeable
		// This allows fine-grained control over which cells can participate in merging
		if cc, exists := t.CellConfigs[colIndex+1]; exists {
			if cellConfig, cellExists := cc[0]; cellExists && !cellConfig.Mergeable {
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
		value, err := item.Lookup(column.Name)
		if err != nil {
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
			shouldMerge := evaluateMergeConditions(lastValue, processedValue, conditions)

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
