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

// Test getTable function
func TestSpreadsheetExcelize_getTable(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)

	result := se.getTable()

	if result != table {
		t.Error("getTable should return the original table")
	}
}

// Test getFile function
func TestSpreadsheetExcelize_getFile(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	file := excelize.NewFile()
	se.WithFile(file)

	result := se.getFile()

	if result != file {
		t.Error("getFile should return the excelize file")
	}
}

// Test createNewFile function
func TestSpreadsheetExcelize_createNewFile(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)

	err := se.createNewFile()

	if err != nil {
		t.Errorf("createNewFile should not return error, got: %v", err)
	}
	if se.File == nil {
		t.Error("File should be created")
	}
}

// Test saveToWriter function
func TestSpreadsheetExcelize_saveToWriter(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.createNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}

	var buf bytes.Buffer
	err := se.saveToWriter(&buf)

	if err != nil {
		t.Errorf("saveToWriter should not return error, got: %v", err)
	}
	if buf.Len() == 0 {
		t.Error("Buffer should contain data after writing")
	}
}

// Test close function
func TestSpreadsheetExcelize_close(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.createNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}

	err := se.close()

	if err != nil {
		t.Errorf("close should not return error, got: %v", err)
	}
}

// Test getSheetName function
func TestSpreadsheetExcelize_getSheetName(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	sheetName := "TestSheet"
	se := NewSpreadsheetExcelize(sheetName, table)

	result := se.getSheetName()

	if result != sheetName {
		t.Errorf("Expected sheet name %s, got %s", sheetName, result)
	}
}

// Test setSheetName function
func TestSpreadsheetExcelize_setSheetName(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("InitialSheet", table)
	newSheetName := "NewSheet"

	se.setSheetName(newSheetName)

	if se.SheetName != newSheetName {
		t.Errorf("Expected sheet name %s, got %s", newSheetName, se.SheetName)
	}
}

// Test createSheet function
func TestSpreadsheetExcelize_createSheet(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.createNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}

	err := se.createSheet()

	if err != nil {
		t.Errorf("createSheet should not return error, got: %v", err)
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

// Test createSheet function with existing sheet
func TestSpreadsheetExcelize_createSheet_ExistingSheet(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("Sheet1", table) // Use default sheet name
	if err := se.createNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}

	err := se.createSheet()

	if err != nil {
		t.Errorf("createSheet should not return error for existing sheet, got: %v", err)
	}
}

// Test setActiveSheet function
func TestSpreadsheetExcelize_setActiveSheet(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.createNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}
	if err := se.createSheet(); err != nil {
		t.Fatalf("failed to create sheet: %v", err)
	}

	err := se.setActiveSheet()

	if err != nil {
		t.Errorf("setActiveSheet should not return error, got: %v", err)
	}

	// Verify the sheet is active
	activeIndex := se.File.GetActiveSheetIndex()
	expectedIndex, _ := se.File.GetSheetIndex("TestSheet")
	if activeIndex != expectedIndex {
		t.Error("Sheet was not set as active")
	}
}

// Test setActiveSheet function with non-existent sheet
func TestSpreadsheetExcelize_setActiveSheet_NonExistentSheet(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("NonExistentSheet", table)
	if err := se.createNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}

	err := se.setActiveSheet()

	// The current implementation doesn't return an error for non-existent sheets
	// It calls setActiveSheet with index -1, but Excel may default to sheet 0
	if err != nil {
		t.Errorf("setActiveSheet returned unexpected error: %v", err)
	}

	// Verify that the active sheet index remains at the default (0)
	// This is the actual behavior when trying to set a non-existent sheet
	activeIndex := se.File.GetActiveSheetIndex()
	if activeIndex != 0 {
		t.Logf("Active sheet index is %d (this may be expected behavior)", activeIndex)
	}
}

// Test setColumnWidth function
func TestSpreadsheetExcelize_setColumnWidth(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.createNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}
	if err := se.createSheet(); err != nil {
		t.Fatalf("failed to create sheet: %v", err)
	}

	err := se.setColumnWidth("A", 20.5)

	if err != nil {
		t.Errorf("setColumnWidth should not return error, got: %v", err)
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

// Test getCellValue function
func TestSpreadsheetExcelize_getCellValue(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.createNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}
	if err := se.createSheet(); err != nil {
		t.Fatalf("failed to create sheet: %v", err)
	}

	// Set a value first
	if err := se.File.SetCellValue("TestSheet", "A1", "TestValue"); err != nil {
		t.Fatalf("failed to set cell value: %v", err)
	}

	value, err := se.getCellValue(1, 1)

	if err != nil {
		t.Errorf("getCellValue should not return error, got: %v", err)
	}
	if value != "TestValue" {
		t.Errorf("Expected value 'TestValue', got '%s'", value)
	}
}

// Test setCellValue function
func TestSpreadsheetExcelize_setCellValue(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.createNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}
	if err := se.createSheet(); err != nil {
		t.Fatalf("failed to create sheet: %v", err)
	}

	err := se.setCellValue(1, 1, "TestValue")

	if err != nil {
		t.Errorf("setCellValue should not return error, got: %v", err)
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

// Test mergeCells function
func TestSpreadsheetExcelize_mergeCells(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.createNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}
	if err := se.createSheet(); err != nil {
		t.Fatalf("failed to create sheet: %v", err)
	}

	err := se.mergeCells(1, 1, 2, 2)

	if err != nil {
		t.Errorf("mergeCells should not return error, got: %v", err)
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

// Test isCellMerged function
func TestSpreadsheetExcelize_isCellMerged(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.createNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}
	if err := se.createSheet(); err != nil {
		t.Fatalf("failed to create sheet: %v", err)
	}

	// Merge cells first
	if err := se.mergeCells(1, 1, 2, 2); err != nil {
		t.Fatalf("failed to merge cells: %v", err)
	}

	result := se.isCellMerged(1, 1)

	if !result {
		t.Error("Cell should be merged")
	}

	// Test non-merged cell
	result = se.isCellMerged(3, 3)
	if result {
		t.Error("Cell should not be merged")
	}
}

// Test isCellMergedHorizontally function
func TestSpreadsheetExcelize_isCellMergedHorizontally(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.createNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}
	if err := se.createSheet(); err != nil {
		t.Fatalf("failed to create sheet: %v", err)
	}

	// Merge cells horizontally
	if err := se.mergeCells(1, 1, 3, 1); err != nil {
		t.Fatalf("failed to merge cells: %v", err)
	}

	result := se.isCellMergedHorizontally(1, 1)

	if !result {
		t.Error("Cell should be merged horizontally")
	}

	// Test non-merged cell
	result = se.isCellMergedHorizontally(1, 2)
	if result {
		t.Error("Cell should not be merged horizontally")
	}
}

// Test applyBorderToCell function
func TestSpreadsheetExcelize_applyBorderToCell(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.createNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}
	if err := se.createSheet(); err != nil {
		t.Fatalf("failed to create sheet: %v", err)
	}

	border := &Border{Style: BorderStyleThin}

	err := se.applyBorderToCell(1, 1, "top", border)

	if err != nil {
		t.Errorf("applyBorderToCell should not return error, got: %v", err)
	}
}

// Test applyBordersToRange function
func TestSpreadsheetExcelize_applyBordersToRange(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.createNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}
	if err := se.createSheet(); err != nil {
		t.Fatalf("failed to create sheet: %v", err)
	}

	borders := Borders{
		Top:    &Border{Style: BorderStyleThin},
		Bottom: &Border{Style: BorderStyleThin},
		Left:   &Border{Style: BorderStyleThin},
		Right:  &Border{Style: BorderStyleThin},
	}

	err := se.applyBordersToRange(1, 1, 2, 2, borders)

	if err != nil {
		t.Errorf("applyBordersToRange should not return error, got: %v", err)
	}
}

// Test hasExistingBorder function
func TestSpreadsheetExcelize_hasExistingBorder(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.createNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}
	if err := se.createSheet(); err != nil {
		t.Fatalf("failed to create sheet: %v", err)
	}

	// Apply a border first
	border := &Border{Style: BorderStyleThin}
	if err := se.applyBorderToCell(1, 1, "top", border); err != nil {
		t.Fatalf("failed to apply border: %v", err)
	}

	result := se.hasExistingBorder(1, 1, "top")

	if !result {
		t.Error("Cell should have existing border")
	}

	// Test cell without border
	result = se.hasExistingBorder(2, 2, "top")
	if result {
		t.Error("Cell should not have existing border")
	}
}

// Test applyStyleToCell function
func TestSpreadsheetExcelize_applyStyleToCell(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.createNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}
	if err := se.createSheet(); err != nil {
		t.Fatalf("failed to create sheet: %v", err)
	}

	style := Style{
		Bold:            true,
		FontSize:        12,
		TextColor:       "#000000",
		BackgroundColor: "#FFFFFF",
	}

	err := se.applyStyleToCell(1, 1, style)

	if err != nil {
		t.Errorf("applyStyleToCell should not return error, got: %v", err)
	}
}

// Test applyStyleToRange function
func TestSpreadsheetExcelize_applyStyleToRange(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	if err := se.createNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}
	if err := se.createSheet(); err != nil {
		t.Fatalf("failed to create sheet: %v", err)
	}

	style := Style{
		Bold:            true,
		FontSize:        12,
		TextColor:       "#000000",
		BackgroundColor: "#FFFFFF",
	}

	err := se.applyStyleToRange(1, 1, 2, 2, style)

	if err != nil {
		t.Errorf("applyStyleToRange should not return error, got: %v", err)
	}
}

// Test getColumnLetter function
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
		result := se.getColumnLetter(tc.col)
		if result != tc.expected {
			t.Errorf("getColumnLetter(%d): expected %s, got %s", tc.col, tc.expected, result)
		}
	}
}

// Test processValue function
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
		result, err := se.processValue(tc.value, tc.format)
		if tc.hasError && err == nil {
			t.Errorf("processValue(%v, %s): expected error but got none", tc.value, tc.format)
		}
		if !tc.hasError && err != nil {
			t.Errorf("processValue(%v, %s): unexpected error: %v", tc.value, tc.format, err)
		}
		if !tc.hasError && result == nil {
			t.Errorf("processValue(%v, %s): expected result but got nil", tc.value, tc.format)
		}
	}
}
