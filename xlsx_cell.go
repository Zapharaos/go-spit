package go_spit

import (
	"fmt"
)

// applyRowConfigs applies row-specific configurations for all configured rows
func (xlsx XLSX) applyCellConfigs() error {
	// Calculate the starting row for data (accounting for header)
	headerShift := xlsx.GetDataStartRow()

	for _, rowConfigs := range xlsx.CellConfigs {
		for _, cellConfig := range rowConfigs {

			// Process each cell configuration
			err := xlsx.applyCellConfig(cellConfig, headerShift)
			if err != nil {
				return fmt.Errorf("failed to apply cell config: %w", err)
			}
		}
	}

	return nil
}

// applyRowConfig applies row-specific configuration (alignment, borders, merging, etc.) to a given row
func (xlsx XLSX) applyCellConfig(config CellConfig, headerShift int) error {
	// Apply borders to the cell
	if config.Border != nil {
		// TODO : Implement cell border application
		/*var err error
		cellRef := fmt.Sprintf("%s%d", xlsx.getColumnLetter(config.ColIndex+1), config.RowIndex+headerShift)

		if err = xlsx.applyCellBorder(cellRef, "left", config.Border.Left); err != nil {
			return fmt.Errorf("failed to apply left cell border: %w", err)
		}
		if err = xlsx.applyCellBorder(cellRef, "right", config.Border.Right); err != nil {
			return fmt.Errorf("failed to apply right cell border: %w", err)
		}
		if err = xlsx.applyCellBorder(cellRef, "top", config.Border.Top); err != nil {
			return fmt.Errorf("failed to apply top cell border: %w", err)
		}
		if err = xlsx.applyCellBorder(cellRef, "bottom", config.Border.Bottom); err != nil {
			return fmt.Errorf("failed to apply bottom cell border: %w", err)
		}*/
	}

	return nil
}
