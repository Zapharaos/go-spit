package go_spit

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

// TODO : refactoring + renaming + using table struct instead of XLSX directly ?

// processColumnsMerging merges cells horizontally using a recursive approach
// This replaces both the old flat and nested merging functions
func (xlsx XLSX) processColumnsMerging() error {
	dataStartRow := xlsx.Table.GetDataStartRow()

	// Get flattened columns since merging applies to leaf columns only
	for actualColIndex, column := range xlsx.Columns.GetFlattenedColumns() {
		if err := xlsx.processColumnVerticalMerging(column, actualColIndex+1, dataStartRow); err != nil {
			L().Warn("Failed to process column for vertical merging", Error(err))
		}
	}

	// Process each row
	for rowIndex, item := range xlsx.Data {
		// Check if row has configured horizontal merging
		if rc, exists := xlsx.RowConfigs[rowIndex]; exists && rc.Merge != nil {
			continue
		}
		rowNum := rowIndex + dataStartRow
		if err := xlsx.processRowHorizontalMergingRecursive(item, xlsx.Columns, rowNum, 1, nil); err != nil {
			L().Warn("Failed to process row for horizontal merging", Int("row", rowNum), Error(err))
		}
	}
	return nil
}

// processColumnVerticalMerging handles the vertical merging logic for a single column
func (xlsx XLSX) processColumnVerticalMerging(column Column, actualColIndex int, dataStartRow int) error {
	mergeConfig := column.Merge
	if mergeConfig == nil || len(mergeConfig.Vertical) == 0 {
		return nil
	}

	// Find merge ranges using unified function
	mergeRanges := xlsx.findVerticalMergeRanges(actualColIndex, column.Name, column.Format, column.Merge.Vertical)

	// Apply merges to the Excel file
	colLetter := xlsx.getColumnLetter(actualColIndex)
	for _, mr := range mergeRanges {
		if len(mr) < 2 {
			continue // Skip single-cell ranges
		}

		startRow := mr[0] + dataStartRow
		endRow := mr[len(mr)-1] + dataStartRow

		startCell := fmt.Sprintf("%s%d", colLetter, startRow)
		endCell := fmt.Sprintf("%s%d", colLetter, endRow)

		if err := xlsx.File.MergeCell(xlsx.SheetName, startCell, endCell); err != nil {
			L().Warn("Failed to merge cells vertically", String("range", fmt.Sprintf("%s:%s", startCell, endCell)), Error(err))
			continue
		}
	}

	return nil
}

// findVerticalMergeRanges identifies ranges of rows that should be merged based on merge conditions
func (xlsx XLSX) findVerticalMergeRanges(colIndex int, fieldName string, format string, conditions []MergeCondition) [][]int {
	var mergeRanges [][]int
	var currentRange []int
	var lastValue interface{}

	for rowIndex, item := range xlsx.Data {

		// Check if row has configured vertical merging
		if rc, exists := xlsx.RowConfigs[rowIndex]; exists && rc.Merge != nil {
			continue
		}

		// Check if cell is mergeable
		if cc, exists := xlsx.CellConfigs[colIndex][rowIndex]; exists && !cc.Mergeable {
			continue
		}

		value, err := item.Lookup(fieldName)
		if err != nil {
			// End current range if we can't get the value
			if len(currentRange) > 1 {
				mergeRanges = append(mergeRanges, currentRange)
			}
			currentRange = nil
			lastValue = nil
			continue
		}

		processedValue, err := xlsx.processValue(value, format)
		if err != nil {
			continue
		}

		if rowIndex == 0 {
			// First row - start new range
			currentRange = []int{rowIndex}
			lastValue = processedValue
		} else {
			// Check if current row should merge with previous using unified function
			shouldMerge := evaluateMergeConditions(lastValue, processedValue, conditions)

			if shouldMerge {
				if len(currentRange) == 0 {
					currentRange = []int{rowIndex - 1, rowIndex}
				} else {
					currentRange = append(currentRange, rowIndex)
				}
			} else {
				// End current range and start new one
				if len(currentRange) > 1 {
					mergeRanges = append(mergeRanges, currentRange)
				}
				currentRange = []int{rowIndex}
			}

			lastValue = processedValue
		}
	}

	// Don't forget the last range
	if len(currentRange) > 1 {
		mergeRanges = append(mergeRanges, currentRange)
	}

	return mergeRanges
}

// processRowHorizontalMergingRecursive processes horizontal merging recursively from deepest level up
func (xlsx XLSX) processRowHorizontalMergingRecursive(item Data, columns Columns, rowNum int, startColIndex int, optRowConfig *RowConfig) error {
	currentColIndex := startColIndex

	// First pass: recursively process all sub-columns (deepest level first)
	for _, column := range columns {
		if column.HasSubColumns() {
			if err := xlsx.processRowHorizontalMergingRecursive(item, column.Columns, rowNum, currentColIndex, optRowConfig); err != nil {
				return err
			}
			currentColIndex += column.GetColumnCount()
		} else {
			currentColIndex++
		}
	}

	// Second pass: merge at current level after sub-columns have been processed
	return xlsx.processLevelHorizontalMerging(item, columns, rowNum, startColIndex, optRowConfig)
}

// processLevelHorizontalMerging handles merging at a specific hierarchical level
func (xlsx XLSX) processLevelHorizontalMerging(item Data, columns Columns, rowNum int, startColIndex int, optRowConfig *RowConfig) error {
	currentColIndex := startColIndex

	for i := 0; i < len(columns); i++ {
		column := columns[i]

		// Check if cell is mergeable
		dataRowIndex := xlsx.GetDataIndexFromRowIndex(rowNum)
		if cc, exists := xlsx.CellConfigs[currentColIndex-1][dataRowIndex]; exists && !cc.Mergeable {
			currentColIndex++
			continue
		}

		// Skip if this column doesn't have merge configuration
		if column.Merge == nil || len(column.Merge.Horizontal) == 0 {
			currentColIndex += column.GetColumnCount()
			continue
		}

		// Find the effective value for this column (considering sub-columns)
		currentValue, err := xlsx.getEffectiveColumnValue(item, column, rowNum, currentColIndex)
		if err != nil {
			return fmt.Errorf("error getting effective value for column %s: %w", column.Name, err)
		}

		// Look for consecutive columns that should merge with this one
		mergeGroup := []columnMerge{{
			column:   column,
			startCol: currentColIndex,
			endCol:   currentColIndex + column.GetColumnCount() - 1,
			value:    currentValue,
		}}

		// Check subsequent columns for merging
		nextColIndex := currentColIndex + column.GetColumnCount()
		for j := i + 1; j < len(columns); j++ {

			// Check if cell is mergeable
			if cc, exists := xlsx.CellConfigs[nextColIndex-1][dataRowIndex]; exists && !cc.Mergeable {
				break
			}

			nextColumn := columns[j]
			nextValue, err := xlsx.getEffectiveColumnValue(item, nextColumn, rowNum, nextColIndex)
			if err != nil {
				return fmt.Errorf("error getting effective value for column %s: %w", nextColumn.Name, err)
			}

			merge := column.Merge
			if optRowConfig != nil && optRowConfig.Merge != nil {
				// Use row-specific merge conditions if available
				merge = optRowConfig.Merge
			}

			// Check if this column should merge with the current group
			shouldMerge := evaluateMergeConditions(currentValue, nextValue, merge.Horizontal)
			if shouldMerge {
				mergeGroup = append(mergeGroup, columnMerge{
					column:   nextColumn,
					startCol: nextColIndex,
					endCol:   nextColIndex + nextColumn.GetColumnCount() - 1,
					value:    nextValue,
				})
				nextColIndex += nextColumn.GetColumnCount()
				i = j // Skip the merged column in the outer loop
			} else {
				break
			}
		}

		// Apply merge if we have multiple columns to merge
		if len(mergeGroup) > 1 {
			if err = xlsx.applyHorizontalMerge(mergeGroup, rowNum); err != nil {
				L().Warn("Failed to apply horizontal merge", Error(err))
			}
		}

		currentColIndex = nextColIndex
	}

	return nil
}

// getEffectiveColumnValue gets the effective value for a column, considering merged sub-columns
func (xlsx XLSX) getEffectiveColumnValue(item Data, column Column, rowNum int, startColIndex int) (string, error) {
	if !column.HasSubColumns() {
		// Leaf column - get value directly
		value, _ := item.Lookup(column.Name)

		// Process the value based on column format
		processedValue, err := xlsx.processValue(value, column.Format)
		if err != nil {
			return "", fmt.Errorf("error processing value for column %s: %w", column.Name)
		}

		return fmt.Sprintf("%v", processedValue), nil
	}

	// Parent column - check if all sub-columns have been merged into one cell
	// If so, get the value from the first cell; otherwise, return empty to prevent merging
	subColumnCount := column.GetColumnCount()
	if subColumnCount == 1 {
		// All sub-columns were merged, get the value from the merged cell
		cellRef, _ := excelize.CoordinatesToCellName(startColIndex, rowNum)
		value, err := xlsx.File.GetCellValue(xlsx.SheetName, cellRef)
		if err != nil {
			return "", err
		}
		return value, nil
	}

	// Sub-columns are not fully merged, so this parent column shouldn't merge
	return "", nil
}

// applyHorizontalMerge applies the actual merge operation for a group of columns
func (xlsx XLSX) applyHorizontalMerge(mergeGroup []columnMerge, rowNum int) error {
	if len(mergeGroup) < 2 {
		return nil
	}

	startCol := mergeGroup[0].startCol
	endCol := mergeGroup[len(mergeGroup)-1].endCol

	return xlsx.applyHorizontalCellMerge(startCol, endCol, rowNum)
}

// applyHorizontalCellMerge is a shared utility function for merging cells horizontally
// This can be used by both column-level and row-level merging functions
func (xlsx XLSX) applyHorizontalCellMerge(startCol, endCol, rowNum int) error {
	if startCol >= endCol {
		return nil
	}

	startCell, _ := excelize.CoordinatesToCellName(startCol, rowNum)
	endCell, _ := excelize.CoordinatesToCellName(endCol, rowNum)

	if err := xlsx.File.MergeCell(xlsx.SheetName, startCell, endCell); err != nil {
		return fmt.Errorf("failed to merge cells %s:%s: %w", startCell, endCell, err)
	}

	return nil
}

// isCellMerged checks if a cell is part of a merged range
func (xlsx XLSX) isCellMerged(cellRef string) bool {
	mergedCells, err := xlsx.File.GetMergeCells(xlsx.SheetName)
	if err != nil {
		return false
	}

	col, row, err := excelize.CellNameToCoordinates(cellRef)
	if err != nil {
		return false
	}

	for _, mergedCell := range mergedCells {
		startCol, startRow, err1 := excelize.CellNameToCoordinates(mergedCell.GetStartAxis())
		endCol, endRow, err2 := excelize.CellNameToCoordinates(mergedCell.GetEndAxis())

		if err1 == nil && err2 == nil {
			if col >= startCol && col <= endCol && row >= startRow && row <= endRow {
				return true
			}
		}
	}
	return false
}

// isCellMergedHorizontally checks if a cell is part of a horizontally merged range (spans multiple columns)
func (xlsx XLSX) isCellMergedHorizontally(cellRef string) bool {
	mergedCells, err := xlsx.File.GetMergeCells(xlsx.SheetName)
	if err != nil {
		return false
	}

	col, row, err := excelize.CellNameToCoordinates(cellRef)
	if err != nil {
		return false
	}

	for _, mergedCell := range mergedCells {
		startCol, startRow, err1 := excelize.CellNameToCoordinates(mergedCell.GetStartAxis())
		endCol, endRow, err2 := excelize.CellNameToCoordinates(mergedCell.GetEndAxis())

		if err1 == nil && err2 == nil {
			if col >= startCol && col <= endCol && row >= startRow && row <= endRow {
				if endCol > startCol { // merged horizontally
					return true
				}
			}
		}
	}
	return false
}

// findUnmergedParentColumn traverses up the column hierarchy to find the first parent column
// whose corresponding cell isn't merged with the current cell
func (xlsx XLSX) findUnmergedParentColumn(colIndex, rowIndex int) *Column {
	// Convert colIndex to 0-based for the GetParentColumnByIndex method
	parentColumn := xlsx.Columns.GetParentColumnByIndex(colIndex - 1)

	// Traverse up the hierarchy until we find an unmerged parent or reach the root
	for parentColumn != nil {
		// Calculate the parent column's cell position
		// For parent columns, we need to find their starting column index
		parentColIndex := xlsx.Columns.GetColumnIndex(parentColumn)
		if parentColIndex > 0 {
			parentCellRef, err := excelize.CoordinatesToCellName(parentColIndex, rowIndex)
			if err == nil && !xlsx.isCellMerged(parentCellRef) {
				return parentColumn
			}
		}

		// Move up to the next parent level
		// Note: This is a simplified approach - in a more complex hierarchy,
		// you might need to implement a more sophisticated parent traversal
		break
	}

	return nil
}
