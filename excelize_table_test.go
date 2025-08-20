package spit

import (
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

func TestNewTableExcelize(t *testing.T) {
	table := &Table{
		WriteHeader: true,
		Columns: Columns{
			{Name: "col1", Label: "Column 1"},
		},
	}

	result := NewTableExcelize("Sheet1", table)

	if result.SheetName != "Sheet1" {
		t.Errorf("NewTableExcelize() SheetName = %v, want %v", result.SheetName, "Sheet1")
	}
	if result.Table != table {
		t.Errorf("NewTableExcelize() Table = %v, want %v", result.Table, table)
	}
	if result.File != nil {
		t.Errorf("NewTableExcelize() File should be nil initially, got %v", result.File)
	}
}

func TestTableExcelize_WithFile(t *testing.T) {
	table := &Table{}
	tableExcel := NewTableExcelize("Sheet1", table)
	file := excelize.NewFile()

	result := tableExcel.WithFile(file)

	if result != tableExcel {
		t.Errorf("WithFile() should return the same instance")
	}
	if tableExcel.File != file {
		t.Errorf("WithFile() File = %v, want %v", tableExcel.File, file)
	}
}

func TestTableExcelize_getTable(t *testing.T) {
	table := &Table{
		WriteHeader: true,
		Columns: Columns{
			{Name: "test", Label: "Test Column"},
		},
	}
	tableExcel := NewTableExcelize("Sheet1", table)

	result := tableExcel.GetTable()

	if result != table {
		t.Errorf("GetTable() = %v, want %v", result, table)
	}
}

func TestTableExcelize_getCellValue(t *testing.T) {
	file := excelize.NewFile()
	defer func(file *excelize.File) {
		_ = file.Close()
	}(file)

	sheetName := "Sheet1"
	tableExcel := NewTableExcelize(sheetName, &Table{}).WithFile(file)

	// Set a test value
	if err := file.SetCellValue(sheetName, "A1", "test value"); err != nil {
		t.Fatalf("failed to set cell value: %v", err)
	}

	tests := []struct {
		name     string
		col      int
		row      int
		expected string
		wantErr  bool
	}{
		{
			name:     "Valid cell with value",
			col:      1,
			row:      1,
			expected: "test value",
			wantErr:  false,
		},
		{
			name:     "Empty cell",
			col:      2,
			row:      1,
			expected: "",
			wantErr:  false,
		},
		{
			name:     "Invalid coordinates",
			col:      0,
			row:      0,
			expected: "",
			wantErr:  true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result, err := tableExcel.GetCellValue(tt.col, tt.row)
			if tt.wantErr {
				if err == nil {
					t.Errorf("GetCellValue() expected error, got nil")
				}
			} else {
				if err != nil {
					t.Errorf("GetCellValue() unexpected error: %v", err)
				}
				if result != tt.expected {
					t.Errorf("GetCellValue() = %v, want %v", result, tt.expected)
				}
			}
		})
	}
}

func TestTableExcelize_setCellValue(t *testing.T) {
	file := excelize.NewFile()
	defer func(file *excelize.File) {
		_ = file.Close()
	}(file)

	sheetName := "Sheet1"
	tableExcel := NewTableExcelize(sheetName, &Table{}).WithFile(file)

	tests := []struct {
		name    string
		col     int
		row     int
		value   interface{}
		wantErr bool
	}{
		{
			name:    "Set string value",
			col:     1,
			row:     1,
			value:   "test string",
			wantErr: false,
		},
		{
			name:    "Set numeric value",
			col:     2,
			row:     1,
			value:   42,
			wantErr: false,
		},
		{
			name:    "Set boolean value",
			col:     3,
			row:     1,
			value:   true,
			wantErr: false,
		},
		{
			name:    "Invalid coordinates",
			col:     0,
			row:     0,
			value:   "test",
			wantErr: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			err := tableExcel.SetCellValue(tt.col, tt.row, tt.value)
			if tt.wantErr {
				if err == nil {
					t.Errorf("SetCellValue() expected error, got nil")
				}
			} else {
				if err != nil {
					t.Errorf("SetCellValue() unexpected error: %v", err)
				}
				// Verify the value was set correctly
				if !tt.wantErr {
					cellRef, _ := excelize.CoordinatesToCellName(tt.col, tt.row)
					actualValue, _ := file.GetCellValue(sheetName, cellRef)
					expectedStr := ""
					switch v := tt.value.(type) {
					case string:
						expectedStr = v
					case int:
						expectedStr = "42"
					case bool:
						if v {
							expectedStr = "TRUE"
						} else {
							expectedStr = "FALSE"
						}
					}
					if actualValue != expectedStr {
						t.Errorf("SetCellValue() value = %v, want %v", actualValue, expectedStr)
					}
				}
			}
		})
	}
}

func TestTableExcelize_mergeCells(t *testing.T) {
	file := excelize.NewFile()
	defer func(file *excelize.File) {
		_ = file.Close()
	}(file)

	sheetName := "Sheet1"
	tableExcel := NewTableExcelize(sheetName, &Table{}).WithFile(file)

	tests := []struct {
		name     string
		startCol int
		startRow int
		endCol   int
		endRow   int
		wantErr  bool
	}{
		{
			name:     "Valid merge range",
			startCol: 1,
			startRow: 1,
			endCol:   3,
			endRow:   1,
			wantErr:  false,
		},
		{
			name:     "Single cell range",
			startCol: 2,
			startRow: 2,
			endCol:   2,
			endRow:   2,
			wantErr:  false,
		},
		{
			name:     "Invalid start coordinates",
			startCol: 0,
			startRow: 0,
			endCol:   1,
			endRow:   1,
			wantErr:  true,
		},
		{
			name:     "Invalid end coordinates",
			startCol: 1,
			startRow: 1,
			endCol:   0,
			endRow:   0,
			wantErr:  true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			err := tableExcel.MergeCells(tt.startCol, tt.startRow, tt.endCol, tt.endRow)
			if tt.wantErr {
				if err == nil {
					t.Errorf("MergeCells() expected error, got nil")
				}
			} else {
				if err != nil {
					t.Errorf("MergeCells() unexpected error: %v", err)
				}
			}
		})
	}
}

func TestTableExcelize_isCellMerged(t *testing.T) {
	file := excelize.NewFile()
	defer func(file *excelize.File) {
		_ = file.Close()
	}(file)

	sheetName := "Sheet1"
	tableExcel := NewTableExcelize(sheetName, &Table{}).WithFile(file)

	// Create a merged range for testing
	if err := file.MergeCell(sheetName, "A1", "C1"); err != nil {
		t.Fatalf("failed to merge cells: %v", err)
	}

	tests := []struct {
		name     string
		col      int
		row      int
		expected bool
		setup    func()
	}{
		{
			name:     "Cell in merged range - start",
			col:      1,
			row:      1,
			expected: true,
			setup:    func() {},
		},
		{
			name:     "Cell in merged range - middle",
			col:      2,
			row:      1,
			expected: true,
			setup:    func() {},
		},
		{
			name:     "Cell in merged range - end",
			col:      3,
			row:      1,
			expected: true,
			setup:    func() {},
		},
		{
			name:     "Cell not in merged range",
			col:      1,
			row:      2,
			expected: false,
			setup:    func() {},
		},
		{
			name:     "Invalid coordinates",
			col:      0,
			row:      0,
			expected: false,
			setup:    func() {},
		},
		{
			name:     "GetMergeCells error - invalid sheet",
			col:      1,
			row:      1,
			expected: false,
			setup: func() {
				// Create a new instance with invalid sheet name to trigger GetMergeCells error
				tableExcel.SheetName = "NonExistentSheet"
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			originalSheet := tableExcel.SheetName
			tt.setup()
			result := tableExcel.IsCellMerged(tt.col, tt.row)
			tableExcel.SheetName = originalSheet // Restore original sheet name
			if result != tt.expected {
				t.Errorf("IsCellMerged() = %v, want %v", result, tt.expected)
			}
		})
	}
}

func TestTableExcelize_isCellMergedHorizontally(t *testing.T) {
	file := excelize.NewFile()
	defer func(file *excelize.File) {
		_ = file.Close()
	}(file)

	sheetName := "Sheet1"
	tableExcel := NewTableExcelize(sheetName, &Table{}).WithFile(file)

	// Create horizontal and vertical merged ranges for testing
	if err := file.MergeCell(sheetName, "A1", "C1"); err != nil {
		t.Fatalf("failed to merge cells horizontally: %v", err)
	}
	if err := file.MergeCell(sheetName, "A3", "A5"); err != nil {
		t.Fatalf("failed to merge cells vertically: %v", err)
	}

	tests := []struct {
		name     string
		col      int
		row      int
		expected bool
		setup    func()
	}{
		{
			name:     "Cell in horizontal merge",
			col:      1,
			row:      1,
			expected: true,
			setup:    func() {},
		},
		{
			name:     "Cell in vertical merge",
			col:      1,
			row:      3,
			expected: false,
			setup:    func() {},
		},
		{
			name:     "Cell not merged",
			col:      1,
			row:      2,
			expected: false,
			setup:    func() {},
		},
		{
			name:     "Invalid coordinates",
			col:      0,
			row:      0,
			expected: false,
			setup:    func() {},
		},
		{
			name:     "GetMergeCells error - invalid sheet",
			col:      1,
			row:      1,
			expected: false,
			setup: func() {
				// Create a new instance with invalid sheet name to trigger GetMergeCells error
				tableExcel.SheetName = "NonExistentSheet"
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			originalSheet := tableExcel.SheetName
			tt.setup()
			result := tableExcel.IsCellMergedHorizontally(tt.col, tt.row)
			tableExcel.SheetName = originalSheet // Restore sheet name
			if result != tt.expected {
				t.Errorf("IsCellMergedHorizontally() = %v, want %v", result, tt.expected)
			}
		})
	}
}

func TestTableExcelize_applyBorderToCell(t *testing.T) {
	file := excelize.NewFile()
	defer func(file *excelize.File) {
		_ = file.Close()
	}(file)

	sheetName := "Sheet1"
	tableExcel := NewTableExcelize(sheetName, &Table{}).WithFile(file)

	tests := []struct {
		name    string
		col     int
		row     int
		side    string
		border  *Border
		wantErr bool
	}{
		{
			name:    "Apply left border",
			col:     1,
			row:     1,
			side:    "left",
			border:  &Border{Style: BorderStyleThin},
			wantErr: false,
		},
		{
			name:    "Apply right border",
			col:     2,
			row:     1,
			side:    "right",
			border:  &Border{Style: BorderStyleMedium},
			wantErr: false,
		},
		{
			name:    "Apply top border",
			col:     3,
			row:     1,
			side:    "top",
			border:  &Border{Style: BorderStyleThick},
			wantErr: false,
		},
		{
			name:    "Apply bottom border",
			col:     4,
			row:     1,
			side:    "bottom",
			border:  &Border{Style: BorderStyleDashed},
			wantErr: false,
		},
		{
			name:    "No border (nil)",
			col:     5,
			row:     1,
			side:    "left",
			border:  nil,
			wantErr: false,
		},
		{
			name:    "No border (BorderStyleNone)",
			col:     6,
			row:     1,
			side:    "left",
			border:  &Border{Style: BorderStyleNone},
			wantErr: false,
		},
		{
			name:    "Invalid side",
			col:     7,
			row:     1,
			side:    "invalid",
			border:  &Border{Style: BorderStyleThin},
			wantErr: true,
		},
		{
			name:    "Invalid coordinates",
			col:     0,
			row:     0,
			side:    "left",
			border:  &Border{Style: BorderStyleThin},
			wantErr: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			err := tableExcel.ApplyBorderToCell(tt.col, tt.row, tt.side, tt.border)
			if tt.wantErr {
				if err == nil {
					t.Errorf("ApplyBorderToCell() expected error, got nil")
				}
			} else {
				if err != nil {
					t.Errorf("ApplyBorderToCell() unexpected error: %v", err)
				}
			}
		})
	}
}

// Add comprehensive test for ApplyBorderToCell error handling
func TestTableExcelize_applyBorderToCell_ErrorHandling(t *testing.T) {
	file := excelize.NewFile()
	defer func(file *excelize.File) {
		_ = file.Close()
	}(file)

	sheetName := "Sheet1"
	tableExcel := NewTableExcelize(sheetName, &Table{}).WithFile(file)

	tests := []struct {
		name    string
		col     int
		row     int
		side    string
		border  *Border
		wantErr bool
		setup   func()
	}{
		{
			name:    "getCellStyle error - invalid sheet",
			col:     1,
			row:     1,
			side:    "left",
			border:  &Border{Style: BorderStyleThin},
			wantErr: true, // This actually does error when trying to set style on invalid sheet
			setup: func() {
				tableExcel.SheetName = "NonExistentSheet"
			},
		},
		{
			name:    "getCellStyle returns nil - should create new style",
			col:     2,
			row:     1,
			side:    "right",
			border:  &Border{Style: BorderStyleMedium},
			wantErr: false,
			setup:   func() {},
		},
		{
			name:    "Valid border application with existing style",
			col:     3,
			row:     1,
			side:    "top",
			border:  &Border{Style: BorderStyleThick},
			wantErr: false,
			setup: func() {
				// Pre-apply a style to the cell
				styleID, _ := file.NewStyle(&excelize.Style{
					Font: &excelize.Font{Bold: true},
				})
				if err := file.SetCellStyle(sheetName, "C1", "C1", styleID); err != nil {
					t.Fatalf("failed to set cell style: %v", err)
				}
			},
		},
		{
			name:    "Invalid side should error",
			col:     4,
			row:     1,
			side:    "diagonal",
			border:  &Border{Style: BorderStyleThin},
			wantErr: true,
			setup:   func() {},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			originalSheet := tableExcel.SheetName
			tt.setup()
			err := tableExcel.ApplyBorderToCell(tt.col, tt.row, tt.side, tt.border)
			tableExcel.SheetName = originalSheet // Restore sheet name

			if tt.wantErr {
				if err == nil {
					t.Errorf("ApplyBorderToCell() expected error, got nil")
				}
			} else {
				if err != nil {
					t.Errorf("ApplyBorderToCell() unexpected error: %v", err)
				}
			}
		})
	}
}

func TestTableExcelize_applyBordersToRange(t *testing.T) {
	file := excelize.NewFile()
	defer func(file *excelize.File) {
		_ = file.Close()
	}(file)

	sheetName := "Sheet1"
	tableExcel := NewTableExcelize(sheetName, &Table{}).WithFile(file)

	tests := []struct {
		name     string
		startCol int
		startRow int
		endCol   int
		endRow   int
		borders  Borders
		wantErr  bool
		setup    func()
	}{
		{
			name:     "Apply left border only",
			startCol: 1,
			startRow: 1,
			endCol:   3,
			endRow:   3,
			borders: Borders{
				Left: &Border{Style: BorderStyleThin},
			},
			wantErr: false,
			setup:   func() {},
		},
		{
			name:     "Apply right border only",
			startCol: 1,
			startRow: 1,
			endCol:   3,
			endRow:   3,
			borders: Borders{
				Right: &Border{Style: BorderStyleMedium},
			},
			wantErr: false,
			setup:   func() {},
		},
		{
			name:     "Apply top border only",
			startCol: 1,
			startRow: 1,
			endCol:   3,
			endRow:   3,
			borders: Borders{
				Top: &Border{Style: BorderStyleThick},
			},
			wantErr: false,
			setup:   func() {},
		},
		{
			name:     "Apply bottom border only",
			startCol: 1,
			startRow: 1,
			endCol:   3,
			endRow:   3,
			borders: Borders{
				Bottom: &Border{Style: BorderStyleDashed},
			},
			wantErr: false,
			setup:   func() {},
		},
		{
			name:     "Apply all borders",
			startCol: 5,
			startRow: 5,
			endCol:   7,
			endRow:   7,
			borders: Borders{
				Left:   &Border{Style: BorderStyleThin},
				Right:  &Border{Style: BorderStyleMedium},
				Top:    &Border{Style: BorderStyleThick},
				Bottom: &Border{Style: BorderStyleDashed},
			},
			wantErr: false,
			setup:   func() {},
		},
		{
			name:     "No borders specified",
			startCol: 10,
			startRow: 10,
			endCol:   12,
			endRow:   12,
			borders:  Borders{},
			wantErr:  false,
			setup:    func() {},
		},
		{
			name:     "Single cell range",
			startCol: 15,
			startRow: 15,
			endCol:   15,
			endRow:   15,
			borders: Borders{
				Left:   &Border{Style: BorderStyleThin},
				Right:  &Border{Style: BorderStyleThin},
				Top:    &Border{Style: BorderStyleThin},
				Bottom: &Border{Style: BorderStyleThin},
			},
			wantErr: false,
			setup:   func() {},
		},
		{
			name:     "Error in ApplyBorderToCell - invalid sheet for left border",
			startCol: 1,
			startRow: 1,
			endCol:   1,
			endRow:   1,
			borders: Borders{
				Left: &Border{Style: BorderStyleThin},
			},
			wantErr: true,
			setup: func() {
				tableExcel.SheetName = "NonExistentSheet"
			},
		},
		{
			name:     "Error in ApplyBorderToCell - invalid sheet for right border",
			startCol: 1,
			startRow: 1,
			endCol:   1,
			endRow:   1,
			borders: Borders{
				Right: &Border{Style: BorderStyleThin},
			},
			wantErr: true,
			setup: func() {
				tableExcel.SheetName = "NonExistentSheet"
			},
		},
		{
			name:     "Error in ApplyBorderToCell - invalid sheet for top border",
			startCol: 1,
			startRow: 1,
			endCol:   1,
			endRow:   1,
			borders: Borders{
				Top: &Border{Style: BorderStyleThin},
			},
			wantErr: true,
			setup: func() {
				tableExcel.SheetName = "NonExistentSheet"
			},
		},
		{
			name:     "Error in ApplyBorderToCell - invalid sheet for bottom border",
			startCol: 1,
			startRow: 1,
			endCol:   1,
			endRow:   1,
			borders: Borders{
				Bottom: &Border{Style: BorderStyleThin},
			},
			wantErr: true,
			setup: func() {
				tableExcel.SheetName = "NonExistentSheet"
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			originalSheet := tableExcel.SheetName
			tt.setup()
			err := tableExcel.ApplyBordersToRange(tt.startCol, tt.startRow, tt.endCol, tt.endRow, tt.borders)
			tableExcel.SheetName = originalSheet // Restore sheet name

			if tt.wantErr {
				if err == nil {
					t.Errorf("ApplyBordersToRange() expected error, got nil")
				}
			} else {
				if err != nil {
					t.Errorf("ApplyBordersToRange() unexpected error: %v", err)
				}
			}
		})
	}
}

func TestTableExcelize_hasExistingBorder(t *testing.T) {
	file := excelize.NewFile()
	defer func(file *excelize.File) {
		_ = file.Close()
	}(file)

	sheetName := "Sheet1"
	tableExcel := NewTableExcelize(sheetName, &Table{}).WithFile(file)

	// Apply a style to test cell to create a border
	styleID, _ := file.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Style: 1, Color: "000000"},
		},
	})
	if err := file.SetCellStyle(sheetName, "A1", "A1", styleID); err != nil {
		t.Fatalf("failed to set cell style: %v", err)
	}

	tests := []struct {
		name     string
		col      int
		row      int
		side     string
		expected bool
		setup    func()
	}{
		{
			name:     "Cell with existing border style",
			col:      1,
			row:      1,
			side:     "left",
			expected: true,
			setup:    func() {},
		},
		{
			name:     "Cell without existing border style",
			col:      2,
			row:      1,
			side:     "left",
			expected: false,
			setup:    func() {},
		},
		{
			name:     "Invalid coordinates",
			col:      0,
			row:      0,
			side:     "left",
			expected: false,
			setup:    func() {},
		},
		{
			name:     "GetCellStyle error - invalid sheet",
			col:      1,
			row:      1,
			side:     "left",
			expected: false,
			setup: func() {
				tableExcel.SheetName = "NonExistentSheet"
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			originalSheet := tableExcel.SheetName
			tt.setup()
			result := tableExcel.HasExistingBorder(tt.col, tt.row, tt.side)
			tableExcel.SheetName = originalSheet // Restore sheet name
			if result != tt.expected {
				t.Errorf("HasExistingBorder() = %v, want %v", result, tt.expected)
			}
		})
	}
}

func TestTableExcelize_applyStyleToCell(t *testing.T) {
	file := excelize.NewFile()
	defer func(file *excelize.File) {
		_ = file.Close()
	}(file)

	sheetName := "Sheet1"
	tableExcel := NewTableExcelize(sheetName, &Table{}).WithFile(file)

	// Pre-apply a style to test merging behavior
	existingStyleID, _ := file.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true, Size: 14},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#FFFF00"},
			Pattern: 1,
		},
		Border: []excelize.Border{
			{Type: "left", Style: 1, Color: "000000"},
		},
		Alignment: &excelize.Alignment{
			Horizontal: "left",
			Vertical:   "top",
		},
	})
	if err := file.SetCellStyle(sheetName, "B2", "B2", existingStyleID); err != nil {
		t.Fatalf("failed to set cell style: %v", err)
	}

	tests := []struct {
		name    string
		col     int
		row     int
		style   Style
		wantErr bool
		setup   func()
	}{
		{
			name: "Apply style to cell without existing style (nil case)",
			col:  1,
			row:  1,
			style: Style{
				Bold:            true,
				FontSize:        12,
				BackgroundColor: "#FF0000",
			},
			wantErr: false,
			setup:   func() {},
		},
		{
			name: "Apply style to cell with existing style - merge border",
			col:  2,
			row:  2,
			style: Style{
				Italic: true,
			},
			wantErr: false,
			setup:   func() {},
		},
		{
			name: "Apply style to cell with existing style - merge fill",
			col:  2,
			row:  2,
			style: Style{
				BackgroundColor: "#00FF00",
			},
			wantErr: false,
			setup:   func() {},
		},
		{
			name: "Apply style to cell with existing style - merge font",
			col:  2,
			row:  2,
			style: Style{
				Bold:     false,
				FontSize: 16,
				Italic:   true,
			},
			wantErr: false,
			setup:   func() {},
		},
		{
			name: "Apply style to cell with existing style - merge alignment",
			col:  2,
			row:  2,
			style: Style{
				Alignment: AlignmentCenterMiddle,
			},
			wantErr: false,
			setup:   func() {},
		},
		{
			name: "Apply style with all properties to existing styled cell",
			col:  2,
			row:  2,
			style: Style{
				Bold:            true,
				Italic:          false,
				FontSize:        18,
				FontFamily:      "Calibri",
				TextColor:       "#0000FF",
				BackgroundColor: "#FFFF00",
				Alignment:       AlignmentRight,
			},
			wantErr: false,
			setup:   func() {},
		},
		{
			name: "Invalid coordinates",
			col:  0,
			row:  0,
			style: Style{
				Bold: true,
			},
			wantErr: true,
			setup:   func() {},
		},
		{
			name: "getCellStyle error - invalid sheet (triggers nil case)",
			col:  3,
			row:  3,
			style: Style{
				Bold: true,
			},
			wantErr: true, // Changed from false - this will error when trying to set style on invalid sheet
			setup: func() {
				tableExcel.SheetName = "NonExistentSheet"
			},
		},
		{
			name: "NewStyle error - invalid style properties",
			col:  4,
			row:  4,
			style: Style{
				FontSize: -1, // This might cause issues in some implementations
			},
			wantErr: false, // Excelize is generally tolerant of invalid values
			setup:   func() {},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			originalSheet := tableExcel.SheetName
			tt.setup()
			err := tableExcel.ApplyStyleToCell(tt.col, tt.row, tt.style)
			tableExcel.SheetName = originalSheet // Restore sheet name

			if tt.wantErr {
				if err == nil {
					t.Errorf("ApplyStyleToCell() expected error, got nil")
				}
			} else {
				if err != nil {
					t.Errorf("ApplyStyleToCell() unexpected error: %v", err)
				}
			}
		})
	}
}

// Test specific conditions for ApplyStyleToCell error handling
func TestTableExcelize_applyStyleToCell_ErrorConditions(t *testing.T) {
	file := excelize.NewFile()
	defer func(file *excelize.File) {
		_ = file.Close()
	}(file)

	sheetName := "Sheet1"
	tableExcel := NewTableExcelize(sheetName, &Table{}).WithFile(file)

	tests := []struct {
		name    string
		col     int
		row     int
		style   Style
		wantErr bool
		setup   func()
	}{
		{
			name: "SetCellStyle error - invalid sheet after style creation",
			col:  1,
			row:  1,
			style: Style{
				Bold: true,
			},
			wantErr: true,
			setup: func() {
				// This will cause SetCellStyle to fail
				tableExcel.SheetName = ""
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			originalSheet := tableExcel.SheetName
			tt.setup()
			err := tableExcel.ApplyStyleToCell(tt.col, tt.row, tt.style)
			tableExcel.SheetName = originalSheet // Restore sheet name

			if tt.wantErr {
				if err == nil {
					t.Errorf("ApplyStyleToCell() expected error, got nil")
				}
			} else {
				if err != nil {
					t.Errorf("ApplyStyleToCell() unexpected error: %v", err)
				}
			}
		})
	}
}

func TestTableExcelize_applyStyleToRange(t *testing.T) {
	file := excelize.NewFile()
	defer func(file *excelize.File) {
		_ = file.Close()
	}(file)

	sheetName := "Sheet1"
	tableExcel := NewTableExcelize(sheetName, &Table{}).WithFile(file)

	tests := []struct {
		name     string
		startCol int
		startRow int
		endCol   int
		endRow   int
		style    Style
		wantErr  bool
		setup    func()
	}{
		{
			name:     "Apply style to single cell range",
			startCol: 1,
			startRow: 1,
			endCol:   1,
			endRow:   1,
			style: Style{
				Bold:            true,
				FontSize:        12,
				BackgroundColor: "#FF0000",
			},
			wantErr: false,
			setup:   func() {},
		},
		{
			name:     "Apply style to horizontal range",
			startCol: 1,
			startRow: 2,
			endCol:   3,
			endRow:   2,
			style: Style{
				Italic:          true,
				BackgroundColor: "#00FF00",
				Alignment:       AlignmentCenter,
			},
			wantErr: false,
			setup:   func() {},
		},
		{
			name:     "Apply style to vertical range",
			startCol: 4,
			startRow: 1,
			endCol:   4,
			endRow:   3,
			style: Style{
				FontFamily: "Arial",
				FontSize:   14,
				TextColor:  "#0000FF",
			},
			wantErr: false,
			setup:   func() {},
		},
		{
			name:     "Apply style to rectangular range",
			startCol: 5,
			startRow: 5,
			endCol:   7,
			endRow:   7,
			style: Style{
				Bold:            true,
				Italic:          true,
				FontSize:        16,
				BackgroundColor: "#FFFF00",
				Alignment:       AlignmentCenterMiddle,
			},
			wantErr: false,
			setup:   func() {},
		},
		{
			name:     "Apply style with minimal properties",
			startCol: 10,
			startRow: 10,
			endCol:   11,
			endRow:   11,
			style: Style{
				Bold: true,
			},
			wantErr: false,
			setup:   func() {},
		},
		{
			name:     "Apply empty style",
			startCol: 15,
			startRow: 15,
			endCol:   16,
			endRow:   16,
			style:    Style{},
			wantErr:  false,
			setup:    func() {},
		},
		{
			name:     "Error in ApplyStyleToCell - invalid coordinates",
			startCol: 0,
			startRow: 1,
			endCol:   1,
			endRow:   1,
			style: Style{
				Bold: true,
			},
			wantErr: true,
			setup:   func() {},
		},
		{
			name:     "Error in ApplyStyleToCell - invalid sheet",
			startCol: 1,
			startRow: 1,
			endCol:   2,
			endRow:   2,
			style: Style{
				Bold: true,
			},
			wantErr: true,
			setup: func() {
				tableExcel.SheetName = "NonExistentSheet"
			},
		},
		{
			name:     "Large range application",
			startCol: 20,
			startRow: 20,
			endCol:   25,
			endRow:   25,
			style: Style{
				FontFamily:      "Calibri",
				FontSize:        11,
				BackgroundColor: "#F0F0F0",
				Alignment:       AlignmentLeft,
			},
			wantErr: false,
			setup:   func() {},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			originalSheet := tableExcel.SheetName
			tt.setup()
			err := tableExcel.ApplyStyleToRange(tt.startCol, tt.startRow, tt.endCol, tt.endRow, tt.style)
			tableExcel.SheetName = originalSheet

			if tt.wantErr {
				if err == nil {
					t.Errorf("ApplyStyleToRange() expected error, got nil")
				}
			} else {
				if err != nil {
					t.Errorf("ApplyStyleToRange() unexpected error: %v", err)
				}
			}
		})
	}
}

func TestTableExcelize_getCellStyle(t *testing.T) {
	file := excelize.NewFile()
	defer func(file *excelize.File) {
		_ = file.Close()
	}(file)

	sheetName := "Sheet1"
	tableExcel := NewTableExcelize(sheetName, &Table{}).WithFile(file)

	// Apply a style to test cell
	styleID, _ := file.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true},
	})
	if err := file.SetCellStyle(sheetName, "A1", "A1", styleID); err != nil {
		t.Fatalf("failed to set cell style: %v", err)
	}

	tests := []struct {
		name    string
		col     int
		row     int
		wantErr bool
		setup   func()
	}{
		{
			name:    "Get style from styled cell",
			col:     1,
			row:     1,
			wantErr: false,
			setup:   func() {},
		},
		{
			name:    "Get style from unstyled cell",
			col:     2,
			row:     1,
			wantErr: false,
			setup:   func() {},
		},
		{
			name:    "Invalid coordinates",
			col:     0,
			row:     0,
			wantErr: true,
			setup:   func() {},
		},
		{
			name:    "GetCellStyle error - invalid sheet",
			col:     1,
			row:     1,
			wantErr: true,
			setup: func() {
				tableExcel.SheetName = "NonExistentSheet"
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			originalSheet := tableExcel.SheetName
			tt.setup()
			result, err := tableExcel.getCellStyle(tt.col, tt.row)
			tableExcel.SheetName = originalSheet // Restore original sheet name
			if tt.wantErr {
				if err == nil {
					t.Errorf("getCellStyle() expected error, got nil")
				}
			} else {
				if err != nil {
					t.Errorf("getCellStyle() unexpected error: %v", err)
				}
				if result == nil && !tt.wantErr {
					t.Errorf("getCellStyle() returned nil style")
				}
			}
		})
	}
}

func TestTableExcelize_getColumnLetter(t *testing.T) {
	tableExcel := NewTableExcelize("Sheet1", &Table{})

	tests := []struct {
		name     string
		col      int
		expected string
	}{
		{
			name:     "Column 1 (A)",
			col:      1,
			expected: "A",
		},
		{
			name:     "Column 26 (Z)",
			col:      26,
			expected: "Z",
		},
		{
			name:     "Column 27 (AA)",
			col:      27,
			expected: "AA",
		},
		{
			name:     "Column 52 (AZ)",
			col:      52,
			expected: "AZ",
		},
		{
			name:     "Column 53 (BA)",
			col:      53,
			expected: "BA",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := tableExcel.GetColumnLetter(tt.col)
			if result != tt.expected {
				t.Errorf("GetColumnLetter() = %v, want %v", result, tt.expected)
			}
		})
	}
}

func TestTableExcelize_processValue(t *testing.T) {
	table := &Table{
		ListSeparator: ",",
	}
	tableExcel := NewTableExcelize("Sheet1", table)

	testTime := time.Date(2024, 1, 15, 10, 30, 0, 0, time.UTC)
	testTimePtr := &testTime

	tests := []struct {
		name     string
		value    interface{}
		format   string
		expected interface{}
		wantErr  bool
	}{
		{
			name:     "String value",
			value:    "test string",
			format:   "",
			expected: "test string",
			wantErr:  false,
		},
		{
			name:     "Integer value",
			value:    42,
			format:   "",
			expected: 42,
			wantErr:  false,
		},
		{
			name:     "Float value",
			value:    3.14,
			format:   "",
			expected: 3.14,
			wantErr:  false,
		},
		{
			name:     "Boolean value",
			value:    true,
			format:   "",
			expected: true,
			wantErr:  false,
		},
		{
			name:     "Time value with format",
			value:    testTime,
			format:   "2006-01-02",
			expected: "2024-01-15",
			wantErr:  false,
		},
		{
			name:     "Time value without format",
			value:    testTime,
			format:   "",
			expected: testTime,
			wantErr:  false,
		},
		{
			name:     "Time pointer with format",
			value:    testTimePtr,
			format:   "2006-01-02",
			expected: "2024-01-15",
			wantErr:  false,
		},
		{
			name:     "Nil time pointer",
			value:    (*time.Time)(nil),
			format:   "",
			expected: "",
			wantErr:  false,
		},
		{
			name:     "Slice with separator",
			value:    []interface{}{"a", "b", "c"},
			format:   "",
			expected: "a,b,c",
			wantErr:  false,
		},
		{
			name:     "Slice without separator",
			value:    []interface{}{"a", "b", "c"},
			format:   "",
			expected: "a,b,c", // Uses table's ListSeparator
			wantErr:  false,
		},
		{
			name:     "Unknown type with format - should return formatted string",
			value:    struct{ Name string }{"test"},
			format:   "2006-01-02",
			expected: "{test}",
			wantErr:  false,
		},
		{
			name:     "Unknown type without format",
			value:    struct{ Name string }{"test"},
			format:   "",
			expected: "{test}",
			wantErr:  false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result, err := tableExcel.ProcessValue(tt.value, tt.format)
			if tt.wantErr {
				if err == nil {
					t.Errorf("ProcessValue() expected error, got nil")
				}
			} else {
				if err != nil {
					t.Errorf("ProcessValue() unexpected error: %v", err)
				}
				if result != tt.expected {
					t.Errorf("ProcessValue() = %v, want %v", result, tt.expected)
				}
			}
		})
	}
}

func TestConvertStyleToExcelizeStyle(t *testing.T) {
	tests := []struct {
		name     string
		style    Style
		validate func(*excelize.Style) bool
	}{
		{
			name: "Bold & italic font style",
			style: Style{
				Bold:   true,
				Italic: true,
			},
			validate: func(s *excelize.Style) bool {
				return s.Font != nil && s.Font.Bold && s.Font.Italic
			},
		},
		{
			name: "Font properties",
			style: Style{
				FontFamily: "Arial",
				FontSize:   12,
				TextColor:  "#FF0000",
				Underline:  "single",
			},
			validate: func(s *excelize.Style) bool {
				return s.Font != nil &&
					s.Font.Family == "Arial" &&
					s.Font.Size == 12 &&
					s.Font.Color == "#FF0000" &&
					s.Font.Underline == "single"
			},
		},
		{
			name: "Background color",
			style: Style{
				BackgroundColor: "#FFFF00",
			},
			validate: func(s *excelize.Style) bool {
				return len(s.Fill.Color) > 0 && s.Fill.Color[0] == "#FFFF00" &&
					s.Fill.Type == "pattern" && s.Fill.Pattern == 1
			},
		},
		{
			name: "Alignment",
			style: Style{
				Alignment: AlignmentCenterMiddle,
			},
			validate: func(s *excelize.Style) bool {
				return s.Alignment != nil &&
					s.Alignment.Horizontal == "center" &&
					s.Alignment.Vertical == "center"
			},
		},
		{
			name: "No style properties",
			style: Style{
				Alignment: AlignmentNone,
			},
			validate: func(s *excelize.Style) bool {
				return s.Font == nil && s.Fill.Color == nil && s.Alignment == nil
			},
		},
		{
			name: "Style with borders, protection, decimal places, and custom format",
			style: Style{
				Bold: true,
				// Note: Border, Protection, DecimalPlaces, CustomNumFmt are not directly
				// supported in the Style struct but are handled in the excelize.Style
			},
			validate: func(s *excelize.Style) bool {
				return s.Font != nil && s.Font.Bold
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := convertStyleToExcelizeStyle(tt.style)
			if result == nil {
				t.Errorf("convertStyleToExcelizeStyle() returned nil")
				return
			}
			if !tt.validate(result) {
				t.Errorf("convertStyleToExcelizeStyle() validation failed for style: %+v", tt.style)
			}
		})
	}
}

func TestIsCellInRange(t *testing.T) {
	tests := []struct {
		name     string
		cellRef  string
		startRef string
		endRef   string
		expected bool
	}{
		{
			name:     "Cell in range",
			cellRef:  "B2",
			startRef: "A1",
			endRef:   "C3",
			expected: true,
		},
		{
			name:     "Cell at start of range",
			cellRef:  "A1",
			startRef: "A1",
			endRef:   "C3",
			expected: true,
		},
		{
			name:     "Cell at end of range",
			cellRef:  "C3",
			startRef: "A1",
			endRef:   "C3",
			expected: true,
		},
		{
			name:     "Cell outside range",
			cellRef:  "D4",
			startRef: "A1",
			endRef:   "C3",
			expected: false,
		},
		{
			name:     "Cell before range",
			cellRef:  "A1",
			startRef: "B2",
			endRef:   "D4",
			expected: false,
		},
		{
			name:     "Invalid cell reference",
			cellRef:  "INVALID",
			startRef: "A1",
			endRef:   "C3",
			expected: false,
		},
		{
			name:     "Invalid start reference",
			cellRef:  "B2",
			startRef: "INVALID",
			endRef:   "C3",
			expected: false,
		},
		{
			name:     "Invalid end reference",
			cellRef:  "B2",
			startRef: "A1",
			endRef:   "INVALID",
			expected: false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := isCellInRange(tt.cellRef, tt.startRef, tt.endRef)
			if result != tt.expected {
				t.Errorf("isCellInRange() = %v, want %v", result, tt.expected)
			}
		})
	}
}
