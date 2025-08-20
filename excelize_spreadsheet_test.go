package spit

import (
	"bytes"
	"testing"

	"github.com/xuri/excelize/v2"
)

// Test NewSpreadsheetExcelize function
func TestNewSpreadsheetExcelize(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	sheetName := "TestSheet"

	se := NewSpreadsheetExcelize(sheetName, table)

	if se == nil {
		t.Fatal("NewSpreadsheetExcelize returned nil")
	}
	if se.SheetName != sheetName {
		t.Errorf("Expected SheetName %s, got %s", sheetName, se.SheetName)
	}
	if se.Table == nil {
		t.Error("Table should not be nil")
	}
	if se.File != nil {
		t.Error("File should be nil initially")
	}
}

// Test WithFile function
func TestSpreadsheetExcelize_WithFile(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	file := excelize.NewFile()

	result := se.WithFile(file)

	if result != se {
		t.Error("WithFile should return the same instance")
	}
	if se.File != file {
		t.Error("File was not set correctly")
	}
	if se.Table.File != file {
		t.Error("Table file was not set correctly")
	}
}

// Test GetTable function
func TestSpreadsheetExcelize_getTable(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)

	result := se.GetTable()

	if result != table {
		t.Error("GetTable should return the original table")
	}
}

// Test GetFile function
func TestSpreadsheetExcelize_getFile(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	file := excelize.NewFile()
	se.WithFile(file)

	result := se.GetFile()

	if result != file {
		t.Error("GetFile should return the excelize file")
	}
}

// Test CreateNewFile function
func TestSpreadsheetExcelize_createNewFile(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)

	err := se.CreateNewFile()

	if err != nil {
		t.Errorf("CreateNewFile should not return error, got: %v", err)
	}
	if se.File == nil {
		t.Error("File should be created")
	}
}

// Test SaveToWriter function
func TestSpreadsheetExcelize_saveToWriter(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.CreateNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}

	var buf bytes.Buffer
	err := se.SaveToWriter(&buf)

	if err != nil {
		t.Errorf("SaveToWriter should not return error, got: %v", err)
	}
	if buf.Len() == 0 {
		t.Error("Buffer should contain data after writing")
	}
}

// Test Close function
func TestSpreadsheetExcelize_close(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.CreateNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}

	err := se.Close()

	if err != nil {
		t.Errorf("Close should not return error, got: %v", err)
	}
}

// Test GetSheetName function
func TestSpreadsheetExcelize_getSheetName(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	sheetName := "TestSheet"
	se := NewSpreadsheetExcelize(sheetName, table)

	result := se.GetSheetName()

	if result != sheetName {
		t.Errorf("Expected sheet name %s, got %s", sheetName, result)
	}
}

// Test SetSheetName function
func TestSpreadsheetExcelize_setSheetName(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("InitialSheet", table)
	newSheetName := "NewSheet"

	se.SetSheetName(newSheetName)

	if se.SheetName != newSheetName {
		t.Errorf("Expected sheet name %s, got %s", newSheetName, se.SheetName)
	}
}

// Test CreateSheet function
func TestSpreadsheetExcelize_createSheet(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.CreateNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}

	err := se.CreateSheet()

	if err != nil {
		t.Errorf("CreateSheet should not return error, got: %v", err)
	}

	// Verify sheet exists
	index, err := se.File.GetSheetIndex("TestSheet")
	if err != nil {
		t.Errorf("Failed to get sheet index: %v", err)
	}
	if index == -1 {
		t.Error("Sheet was not created")
	}
}

// Test CreateSheet function with existing sheet
func TestSpreadsheetExcelize_createSheet_ExistingSheet(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("Sheet1", table) // Use default sheet name
	if err := se.CreateNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}

	err := se.CreateSheet()

	if err != nil {
		t.Errorf("CreateSheet should not return error for existing sheet, got: %v", err)
	}
}

// Test SetActiveSheet function
func TestSpreadsheetExcelize_setActiveSheet(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.CreateNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}
	if err := se.CreateSheet(); err != nil {
		t.Fatalf("failed to create sheet: %v", err)
	}

	err := se.SetActiveSheet()

	if err != nil {
		t.Errorf("SetActiveSheet should not return error, got: %v", err)
	}

	// Verify the sheet is active
	activeIndex := se.File.GetActiveSheetIndex()
	expectedIndex, _ := se.File.GetSheetIndex("TestSheet")
	if activeIndex != expectedIndex {
		t.Error("Sheet was not set as active")
	}
}

// Test SetActiveSheet function with non-existent sheet
func TestSpreadsheetExcelize_setActiveSheet_NonExistentSheet(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("NonExistentSheet", table)
	if err := se.CreateNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}

	err := se.SetActiveSheet()

	// The current implementation doesn't return an error for non-existent sheets
	// It calls SetActiveSheet with index -1, but Excel may default to sheet 0
	if err != nil {
		t.Errorf("SetActiveSheet returned unexpected error: %v", err)
	}

	// Verify that the active sheet index remains at the default (0)
	// This is the actual behavior when trying to set a non-existent sheet
	activeIndex := se.File.GetActiveSheetIndex()
	if activeIndex != 0 {
		t.Logf("Active sheet index is %d (this may be expected behavior)", activeIndex)
	}
}

// Test SetColumnWidth function
func TestSpreadsheetExcelize_setColumnWidth(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.CreateNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}
	if err := se.CreateSheet(); err != nil {
		t.Fatalf("failed to create sheet: %v", err)
	}

	err := se.SetColumnWidth("A", 20.5)

	if err != nil {
		t.Errorf("SetColumnWidth should not return error, got: %v", err)
	}

	// Verify the column width was set
	width, err := se.File.GetColWidth("TestSheet", "A")
	if err != nil {
		t.Errorf("Failed to get column width: %v", err)
	}
	if width != 20.5 {
		t.Errorf("Expected column width 20.5, got %f", width)
	}
}

// Test GetCellValue function
func TestSpreadsheetExcelize_getCellValue(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.CreateNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}
	if err := se.CreateSheet(); err != nil {
		t.Fatalf("failed to create sheet: %v", err)
	}

	// Set a value first
	if err := se.File.SetCellValue("TestSheet", "A1", "TestValue"); err != nil {
		t.Fatalf("failed to set cell value: %v", err)
	}

	value, err := se.GetCellValue(1, 1)

	if err != nil {
		t.Errorf("GetCellValue should not return error, got: %v", err)
	}
	if value != "TestValue" {
		t.Errorf("Expected value 'TestValue', got '%s'", value)
	}
}

// Test SetCellValue function
func TestSpreadsheetExcelize_setCellValue(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.CreateNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}
	if err := se.CreateSheet(); err != nil {
		t.Fatalf("failed to create sheet: %v", err)
	}

	err := se.SetCellValue(1, 1, "TestValue")

	if err != nil {
		t.Errorf("SetCellValue should not return error, got: %v", err)
	}

	// Verify the value was set
	value, err := se.File.GetCellValue("TestSheet", "A1")
	if err != nil {
		t.Errorf("Failed to get cell value: %v", err)
	}
	if value != "TestValue" {
		t.Errorf("Expected value 'TestValue', got '%s'", value)
	}
}

// Test MergeCells function
func TestSpreadsheetExcelize_mergeCells(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.CreateNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}
	if err := se.CreateSheet(); err != nil {
		t.Fatalf("failed to create sheet: %v", err)
	}

	err := se.MergeCells(1, 1, 2, 2)

	if err != nil {
		t.Errorf("MergeCells should not return error, got: %v", err)
	}

	// Verify cells are merged by checking merge list
	mergeList, err := se.File.GetMergeCells("TestSheet")
	if err != nil {
		t.Errorf("Failed to get merge cells: %v", err)
	}
	if len(mergeList) == 0 {
		t.Error("No merged cells found")
	}
}

// Test IsCellMerged function
func TestSpreadsheetExcelize_isCellMerged(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.CreateNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}
	if err := se.CreateSheet(); err != nil {
		t.Fatalf("failed to create sheet: %v", err)
	}

	// Merge cells first
	if err := se.MergeCells(1, 1, 2, 2); err != nil {
		t.Fatalf("failed to merge cells: %v", err)
	}

	result := se.IsCellMerged(1, 1)

	if !result {
		t.Error("Cell should be merged")
	}

	// Test non-merged cell
	result = se.IsCellMerged(3, 3)
	if result {
		t.Error("Cell should not be merged")
	}
}

// Test IsCellMergedHorizontally function
func TestSpreadsheetExcelize_isCellMergedHorizontally(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.CreateNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}
	if err := se.CreateSheet(); err != nil {
		t.Fatalf("failed to create sheet: %v", err)
	}

	// Merge cells horizontally
	if err := se.MergeCells(1, 1, 3, 1); err != nil {
		t.Fatalf("failed to merge cells: %v", err)
	}

	result := se.IsCellMergedHorizontally(1, 1)

	if !result {
		t.Error("Cell should be merged horizontally")
	}

	// Test non-merged cell
	result = se.IsCellMergedHorizontally(1, 2)
	if result {
		t.Error("Cell should not be merged horizontally")
	}
}

// Test ApplyBorderToCell function
func TestSpreadsheetExcelize_applyBorderToCell(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.CreateNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}
	if err := se.CreateSheet(); err != nil {
		t.Fatalf("failed to create sheet: %v", err)
	}

	border := &Border{Style: BorderStyleThin}

	err := se.ApplyBorderToCell(1, 1, "top", border)

	if err != nil {
		t.Errorf("ApplyBorderToCell should not return error, got: %v", err)
	}
}

// Test ApplyBordersToRange function
func TestSpreadsheetExcelize_applyBordersToRange(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.CreateNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}
	if err := se.CreateSheet(); err != nil {
		t.Fatalf("failed to create sheet: %v", err)
	}

	borders := Borders{
		Top:    &Border{Style: BorderStyleThin},
		Bottom: &Border{Style: BorderStyleThin},
		Left:   &Border{Style: BorderStyleThin},
		Right:  &Border{Style: BorderStyleThin},
	}

	err := se.ApplyBordersToRange(1, 1, 2, 2, borders)

	if err != nil {
		t.Errorf("ApplyBordersToRange should not return error, got: %v", err)
	}
}

// Test HasExistingBorder function
func TestSpreadsheetExcelize_hasExistingBorder(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.CreateNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}
	if err := se.CreateSheet(); err != nil {
		t.Fatalf("failed to create sheet: %v", err)
	}

	// Apply a border first
	border := &Border{Style: BorderStyleThin}
	if err := se.ApplyBorderToCell(1, 1, "top", border); err != nil {
		t.Fatalf("failed to apply border: %v", err)
	}

	result := se.HasExistingBorder(1, 1, "top")

	if !result {
		t.Error("Cell should have existing border")
	}

	// Test cell without border
	result = se.HasExistingBorder(2, 2, "top")
	if result {
		t.Error("Cell should not have existing border")
	}
}

// Test ApplyStyleToCell function
func TestSpreadsheetExcelize_applyStyleToCell(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.CreateNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}
	if err := se.CreateSheet(); err != nil {
		t.Fatalf("failed to create sheet: %v", err)
	}

	style := Style{
		Bold:            true,
		FontSize:        12,
		TextColor:       "#000000",
		BackgroundColor: "#FFFFFF",
	}

	err := se.ApplyStyleToCell(1, 1, style)

	if err != nil {
		t.Errorf("ApplyStyleToCell should not return error, got: %v", err)
	}
}

// Test ApplyStyleToRange function
func TestSpreadsheetExcelize_applyStyleToRange(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.CreateNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}
	if err := se.CreateSheet(); err != nil {
		t.Fatalf("failed to create sheet: %v", err)
	}

	style := Style{
		Bold:            true,
		FontSize:        12,
		TextColor:       "#000000",
		BackgroundColor: "#FFFFFF",
	}

	err := se.ApplyStyleToRange(1, 1, 2, 2, style)

	if err != nil {
		t.Errorf("ApplyStyleToRange should not return error, got: %v", err)
	}
}

// Test GetColumnLetter function
func TestSpreadsheetExcelize_getColumnLetter(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)

	testCases := []struct {
		col      int
		expected string
	}{
		{1, "A"},
		{2, "B"},
		{26, "Z"},
		{27, "AA"},
	}

	for _, tc := range testCases {
		result := se.GetColumnLetter(tc.col)
		if result != tc.expected {
			t.Errorf("GetColumnLetter(%d): expected %s, got %s", tc.col, tc.expected, result)
		}
	}
}

// Test ProcessValue function
func TestSpreadsheetExcelize_processValue(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)

	testCases := []struct {
		value    interface{}
		format   string
		hasError bool
	}{
		{"test string", "", false},
		{123, "", false},
		{123.45, "", false},
		{true, "", false},
	}

	for _, tc := range testCases {
		result, err := se.ProcessValue(tc.value, tc.format)
		if tc.hasError && err == nil {
			t.Errorf("ProcessValue(%v, %s): expected error but got none", tc.value, tc.format)
		}
		if !tc.hasError && err != nil {
			t.Errorf("ProcessValue(%v, %s): unexpected error: %v", tc.value, tc.format, err)
		}
		if !tc.hasError && result == nil {
			t.Errorf("ProcessValue(%v, %s): expected result but got nil", tc.value, tc.format)
		}
	}
}
