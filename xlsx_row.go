package go_spit

import (
	"fmt"
)

// applyRowConfigs applies row-specific configurations for all configured rows
func (xlsx XLSX) applyRowConfigs() error {
	// Calculate the starting row for data (accounting for header)
	headerShift := xlsx.getDataStartRow()
	totalRows := headerShift + len(xlsx.Data) - 1

	totalColumns := xlsx.Columns.getTotalColumnCount()

	for i := headerShift; i <= totalRows; i++ {
		if err := xlsx.applyRowConfig(i, totalColumns); err != nil {
			return fmt.Errorf("failed to apply row config for row %d: %w", i, err)
		}
	}
	return nil
}

// applyRowConfig applies row-specific configuration (alignment, borders, merging, etc.) to a given row
func (xlsx XLSX) applyRowConfig(rowIndex, totalColumns int) error {
	config, exists := xlsx.RowConfigs[xlsx.getDataIndexFromRowIndex(rowIndex)] // Get config for this row
	if !exists {
		return nil
	}

	// Apply borders to the row
	if config.Border.hasBorders() {
		// TODO : Implement
		/*if err := xlsx.applyRowBorders(rowIndex, totalColumns, *config.Border); err != nil {
			return fmt.Errorf("failed to apply row borders: %w", err)
		}*/
	}

	return nil
}
