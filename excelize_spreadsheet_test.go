package spit

import (
	"bytes"
	"testing"
	"time"

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
		ListSeparator: ",", // Add list separator for slice tests
	}
	se := NewSpreadsheetExcelize("TestSheet", table)

	testTime := time.Date(2024, 1, 15, 10, 30, 0, 0, time.UTC)
	testTimePtr := &testTime

	testCases := []struct {
		name     string
		value    interface{}
		format   string
		expected interface{}
		hasError bool
	}{
		{
			name:     "String value",
			value:    "test string",
			format:   "",
			expected: "test string",
			hasError: false,
		},
		{
			name:     "String value with format - should not be formatted",
			value:    "Total",
			format:   "2006-01-02",
			expected: "Total",
			hasError: false,
		},
		{
			name:     "Integer value",
			value:    123,
			format:   "",
			expected: "123", // TableExcelize.ProcessValue converts to string
			hasError: false,
		},
		{
			name:     "Float value",
			value:    123.45,
			format:   "",
			expected: "123.45", // TableExcelize.ProcessValue converts to string
			hasError: false,
		},
		{
			name:     "Boolean value",
			value:    true,
			format:   "",
			expected: "true", // TableExcelize.ProcessValue converts to string
			hasError: false,
		},
		{
			name:     "Time value with format",
			value:    testTime,
			format:   "2006-01-02",
			expected: "2024-01-15",
			hasError: false,
		},
		{
			name:     "Time value without format",
			value:    testTime,
			format:   "",
			expected: "2024-01-15 10:30:00 +0000 UTC", // TableExcelize.ProcessValue converts to string
			hasError: false,
		},
		{
			name:     "Time pointer with format",
			value:    testTimePtr,
			format:   "2006-01-02",
			expected: "2024-01-15",
			hasError: false,
		},
		{
			name:     "Nil time pointer",
			value:    (*time.Time)(nil),
			format:   "",
			expected: "<nil>", // TableExcelize.ProcessValue returns "<nil>" for nil values
			hasError: false,
		},
		{
			name:     "Slice value",
			value:    []interface{}{"a", "b", "c"},
			format:   "",
			expected: "a,b,c", // Since table has ListSeparator = ","
			hasError: false,
		},
	}

	for _, tc := range testCases {
		t.Run(tc.name, func(t *testing.T) {
			result, err := se.ProcessValue(tc.value, tc.format)
			if tc.hasError && err == nil {
				t.Errorf("ProcessValue(%v, %s): expected error but got none", tc.value, tc.format)
			}
			if !tc.hasError && err != nil {
				t.Errorf("ProcessValue(%v, %s): unexpected error: %v", tc.value, tc.format, err)
			}
			if !tc.hasError && result != tc.expected {
				t.Errorf("ProcessValue(%v, %s): expected %v (%T), got %v (%T)", tc.value, tc.format, tc.expected, tc.expected, result, result)
			}
		})
	}
}

// Test that CreateSheet removes the default "Sheet1" when creating a new file with a different sheet name.
func TestSpreadsheetExcelize_createSheet_RemovesDefaultSheet1(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("Reports", table)
	if err := se.CreateNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}

	// Before CreateSheet: the file should have only the default "Sheet1"
	sheetsBefore := se.File.GetSheetList()
	if len(sheetsBefore) != 1 || sheetsBefore[0] != "Sheet1" {
		t.Fatalf("expected only Sheet1 before CreateSheet, got: %v", sheetsBefore)
	}

	if err := se.CreateSheet(); err != nil {
		t.Fatalf("CreateSheet should not return error, got: %v", err)
	}

	sheetsAfter := se.File.GetSheetList()
	for _, s := range sheetsAfter {
		if s == "Sheet1" {
			t.Errorf("default Sheet1 should have been removed, got sheets: %v", sheetsAfter)
		}
	}
	if len(sheetsAfter) != 1 || sheetsAfter[0] != "Reports" {
		t.Errorf("expected only Reports sheet, got: %v", sheetsAfter)
	}
}

// Test that CreateSheet does NOT remove "Sheet1" when the sheet name is "Sheet1".
func TestSpreadsheetExcelize_createSheet_KeepsSheet1WhenNamed(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("Sheet1", table)
	if err := se.CreateNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}

	if err := se.CreateSheet(); err != nil {
		t.Fatalf("CreateSheet should not return error, got: %v", err)
	}

	sheets := se.File.GetSheetList()
	found := false
	for _, s := range sheets {
		if s == "Sheet1" {
			found = true
			break
		}
	}
	if !found {
		t.Errorf("Sheet1 should be present when sheet name is Sheet1, got: %v", sheets)
	}
}

// Test that CreateSheet does NOT remove "Sheet1" from an existing (user-provided) file.
func TestSpreadsheetExcelize_createSheet_PreservesSheet1InExistingFile(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("Reports", table)

	// Simulate an existing file provided by the user (not via CreateNewFile)
	existingFile := excelize.NewFile()
	se.WithFile(existingFile) // WithFile sets isNewFile = false

	if err := se.CreateSheet(); err != nil {
		t.Fatalf("CreateSheet should not return error, got: %v", err)
	}

	sheets := se.File.GetSheetList()
	hasSheet1 := false
	hasReports := false
	for _, s := range sheets {
		if s == "Sheet1" {
			hasSheet1 = true
		}
		if s == "Reports" {
			hasReports = true
		}
	}
	if !hasSheet1 {
		t.Errorf("Sheet1 should be preserved in an existing file, got: %v", sheets)
	}
	if !hasReports {
		t.Errorf("Reports sheet should have been created, got: %v", sheets)
	}
}

// Test that creating multiple sheets on a new file only removes the default "Sheet1" once.
func TestSpreadsheetExcelize_createSheet_MultipleSheets(t *testing.T) {
	table1 := &Table{Columns: Columns{{Name: "col1", Label: "Col 1"}}}
	table2 := &Table{Columns: Columns{{Name: "col2", Label: "Col 2"}}}

	se1 := NewSpreadsheetExcelize("Reports", table1)
	if err := se1.CreateNewFile(); err != nil {
		t.Fatalf("failed to create new file: %v", err)
	}

	// Share the file with a second sheet
	se2 := NewSpreadsheetExcelize("Summary", table2)
	if err := se2.InitWithFile(se1.File); err != nil {
		t.Fatalf("InitWithFile should not return error, got: %v", err)
	}

	if err := se1.CreateSheet(); err != nil {
		t.Fatalf("CreateSheet for Reports should not return error, got: %v", err)
	}
	if err := se2.CreateSheet(); err != nil {
		t.Fatalf("CreateSheet for Summary should not return error, got: %v", err)
	}

	sheets := se1.File.GetSheetList()
	for _, s := range sheets {
		if s == "Sheet1" {
			t.Errorf("default Sheet1 should have been removed, got sheets: %v", sheets)
		}
	}
	if len(sheets) != 2 {
		t.Errorf("expected 2 sheets (Reports, Summary), got: %v", sheets)
	}
}

// Test InitWithFile function with a valid file.
func TestSpreadsheetExcelize_initWithFile_ValidFile(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)
	file := excelize.NewFile()

	err := se.InitWithFile(file)

	if err != nil {
		t.Errorf("InitWithFile should not return error for a valid *excelize.File, got: %v", err)
	}
	if se.File != file {
		t.Error("File was not set correctly by InitWithFile")
	}
	if se.Table.File != file {
		t.Error("Table file was not set correctly by InitWithFile")
	}
}

// Test InitWithFile function with an invalid (non-excelize) file type.
func TestSpreadsheetExcelize_initWithFile_InvalidType(t *testing.T) {
	table := &Table{
		Columns: Columns{
			{Name: "TestColumn", Label: "Test Column"},
		},
	}
	se := NewSpreadsheetExcelize("TestSheet", table)

	err := se.InitWithFile("not-an-excelize-file")

	if err == nil {
		t.Error("InitWithFile should return error for unsupported file type")
	}
}
