package go_spit

import "fmt"

// applyRowConfigs applies row-specific configurations for all configured rows
func (xlsx XLSX) applyRowConfigs() error {
	// Calculate the starting row for data (accounting for header)
	headerShift := xlsx.GetDataStartRow()
	totalRows := headerShift + len(xlsx.Data) - 1

	totalColumns := xlsx.Columns.GetTotalColumnCount()

	for i := headerShift; i <= totalRows; i++ {
		if err := xlsx.applyRowConfig(i, totalColumns); err != nil {
			return fmt.Errorf("failed to apply row config for row %d: %w", i, err)
		}
	}
	return nil
}

// applyRowConfig applies row-specific configuration (alignment, borders, merging, etc.) to a given row
func (xlsx XLSX) applyRowConfig(rowIndex, totalColumns int) error {
	config, exists := xlsx.RowConfigs[xlsx.GetDataIndexFromRowIndex(rowIndex)] // Get config for this row
	if !exists {
		return nil
	}

	// Apply horizontal merging if configured
	if config.Merge != nil && len(config.Merge.Horizontal) > 0 {
		if err := xlsx.applyRowHorizontalMerging(rowIndex, &config); err != nil {
			return fmt.Errorf("failed to apply row horizontal merging: %w", err)
		}
	}

	// Apply borders to the row
	if config.Border.HasBorders() {
		// TODO : Implement
		/*if err := xlsx.applyRowBorders(rowIndex, totalColumns, *config.Border); err != nil {
			return fmt.Errorf("failed to apply row borders: %w", err)
		}*/
	}

	return nil
}

// applyRowHorizontalMerging merges cells horizontally across an entire row based on merge conditions
func (xlsx XLSX) applyRowHorizontalMerging(rowIndex int, rowConfig *RowConfig) error {
	dataRowIndex := xlsx.GetDataIndexFromRowIndex(rowIndex)
	item := xlsx.Data[dataRowIndex]

	return xlsx.processRowHorizontalMergingRecursive(item, xlsx.Columns, rowIndex, 1, rowConfig)
}
