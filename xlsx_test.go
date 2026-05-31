package spit

import (
	"errors"
	"os"
	"testing"

	"go.uber.org/mock/gomock"
)

// TestExportXLSX tests the main ExportXLSX function
func TestExportXLSX(t *testing.T) {
	ctrl := gomock.NewController(t)
	defer ctrl.Finish()

	// Create a temporary directory for test files
	tempDir, err := os.MkdirTemp("", "xlsx_test")
	if err != nil {
		t.Fatalf("Failed to create temp directory: %v", err)
	}
	defer func() {
		if err := os.RemoveAll(tempDir); err != nil {
			t.Logf("Failed to remove temp directory: %v", err)
		}
	}()

	tests := []struct {
		name           string
		setupMock      func(*MockSpreadsheet)
		params         FileWriteParams
		expectError    bool
		errorContains  string
		validateResult func(*FileWriteResult)
	}{
		{
			name: "successful_export_with_new_file",
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().GetFile().Return(nil)
				mock.EXPECT().CreateNewFile().Return(nil)
				mock.EXPECT().Close().Return(nil)
				mock.EXPECT().GetSheetName().Return("")
				mock.EXPECT().SetSheetName("Sheet1")
				mock.EXPECT().CreateSheet().Return(nil)
				mock.EXPECT().SetActiveSheet().Return(nil)
				mock.EXPECT().GetTable().Return(&Table{
					Data: DataSlice{
						{"name": "John", "age": 30},
					},
					Columns: Columns{
						{Name: "name", Label: "Name"},
						{Name: "age", Label: "Age"},
					},
					WriteHeader: true,
				}).AnyTimes()

				// Header writing expectations
				mock.EXPECT().SetCellValue(1, 1, "Name").Return(nil)
				mock.EXPECT().SetCellValue(2, 1, "Age").Return(nil)

				// Data writing expectations
				mock.EXPECT().ProcessValue("John", "").Return("John", nil)
				mock.EXPECT().SetCellValue(1, 2, "John").Return(nil)
				mock.EXPECT().ProcessValue(30, "").Return(30, nil)
				mock.EXPECT().SetCellValue(2, 2, 30).Return(nil)

				// Auto-fit columns expectations
				mock.EXPECT().GetColumnLetter(1).Return("A")
				mock.EXPECT().SetColumnWidth("A", 15.0).Return(nil)
				mock.EXPECT().GetColumnLetter(2).Return("B")
				mock.EXPECT().SetColumnWidth("B", 15.0).Return(nil)

				// Table processing expectations (merging and styling are called on the table, not the spreadsheet)
				// We need to mock any calls that these methods might make to the spreadsheet
				mock.EXPECT().ApplyBorderToCell(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().ApplyBordersToRange(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().ApplyStyleToCell(gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().ApplyStyleToRange(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().MergeCells(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().IsCellMerged(gomock.Any(), gomock.Any()).Return(false).AnyTimes()
				mock.EXPECT().IsCellMergedHorizontally(gomock.Any(), gomock.Any()).Return(false).AnyTimes()
				mock.EXPECT().HasExistingBorder(gomock.Any(), gomock.Any(), gomock.Any()).Return(false).AnyTimes()
				mock.EXPECT().GetCellValue(gomock.Any(), gomock.Any()).Return("", nil).AnyTimes()
				mock.EXPECT().SetCellValue(gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().GetColumnLetter(gomock.Any()).Return("A").AnyTimes()
				mock.EXPECT().ProcessValue(gomock.Any(), gomock.Any()).Return(gomock.Any(), nil).AnyTimes()

				// Final save expectation
				mock.EXPECT().SaveToWriter(gomock.Any()).Return(nil)
			},
			params: FileWriteParams{
				Filename: "test_export",
				Filepath: tempDir,
			},
			expectError: false,
			validateResult: func(result *FileWriteResult) {
				if result == nil {
					t.Error("Expected non-nil result")
					return
				}
				if result.Filename != "test_export.xlsx" {
					t.Errorf("Expected filename 'test_export.xlsx', got %s", result.Filename)
				}
			},
		},
		{
			name: "successful_export_with_existing_file",
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().GetFile().Return(&struct{ name string }{name: "existing_file"})
				mock.EXPECT().GetSheetName().Return("Sheet1")
				mock.EXPECT().CreateSheet().Return(nil)
				mock.EXPECT().SetActiveSheet().Return(nil)
				mock.EXPECT().GetTable().Return(&Table{
					Data:        DataSlice{},
					Columns:     Columns{},
					WriteHeader: false,
				}).AnyTimes()
				mock.EXPECT().SaveToWriter(gomock.Any()).Return(nil)
			},
			params: FileWriteParams{
				Filename: "test_existing",
				Filepath: tempDir,
			},
			expectError: false,
		},
		{
			name: "error_creating_new_file",
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().GetFile().Return(nil)
				mock.EXPECT().CreateNewFile().Return(errors.New("create file error"))
			},
			params: FileWriteParams{
				Filename: "test_error",
				Filepath: tempDir,
			},
			expectError:   true,
			errorContains: "failed to create new XLSX file",
		},
		{
			name: "error_writing_data",
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().GetFile().Return(&struct{ name string }{name: "existing_file"})
				mock.EXPECT().GetSheetName().Return("Sheet1")
				mock.EXPECT().CreateSheet().Return(errors.New("create sheet error"))
			},
			params: FileWriteParams{
				Filename: "test_write_error",
				Filepath: tempDir,
			},
			expectError:   true,
			errorContains: "failed to write data to XLSX file",
		},
		{
			name: "error_saving_to_writer",
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().GetFile().Return(&struct{ name string }{name: "existing_file"})
				mock.EXPECT().GetSheetName().Return("Sheet1")
				mock.EXPECT().CreateSheet().Return(nil)
				mock.EXPECT().SetActiveSheet().Return(nil)
				mock.EXPECT().GetTable().Return(&Table{
					Data:        DataSlice{},
					Columns:     Columns{},
					WriteHeader: false,
				}).AnyTimes()
				mock.EXPECT().SaveToWriter(gomock.Any()).Return(errors.New("save error"))
			},
			params: FileWriteParams{
				Filename: "test_save_error",
				Filepath: tempDir,
			},
			expectError:   true,
			errorContains: "failed to write XLSX to writer",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			mockSpreadsheet := NewMockSpreadsheet(ctrl)
			tt.setupMock(mockSpreadsheet)

			result, err := ExportXLSX(mockSpreadsheet, tt.params)

			if tt.expectError {
				if err == nil {
					t.Error("Expected error but got none")
				} else if tt.errorContains != "" && !contains(err.Error(), tt.errorContains) {
					t.Errorf("Expected error to contain '%s', got '%s'", tt.errorContains, err.Error())
				}
			} else {
				if err != nil {
					t.Errorf("Unexpected error: %v", err)
				}
				if tt.validateResult != nil {
					tt.validateResult(result)
				}
			}
		})
	}
}

// TestXlsx_writeData tests the writeData method
func TestXlsx_writeData(t *testing.T) {
	tests := []struct {
		name        string
		xlsx        *xlsx
		setupMock   func(*MockSpreadsheet)
		expectError bool
		errorMsg    string
	}{
		{
			name: "successful_write_data_with_headers",
			xlsx: &xlsx{
				params: FileWriteParams{},
			},
			setupMock: func(mock *MockSpreadsheet) {
				table := &Table{
					Data: DataSlice{
						{"name": "John", "age": 30},
						{"name": "Jane", "age": 25},
					},
					Columns: Columns{
						{Name: "name", Label: "Name"},
						{Name: "age", Label: "Age"},
					},
				}

				mock.EXPECT().GetSheetName().Return("")
				mock.EXPECT().SetSheetName("Sheet1")
				mock.EXPECT().CreateSheet().Return(nil)
				mock.EXPECT().SetActiveSheet().Return(nil)
				mock.EXPECT().GetTable().Return(table).AnyTimes()

				// Header writing expectations
				mock.EXPECT().SetCellValue(1, 1, "Name").Return(nil)
				mock.EXPECT().SetCellValue(2, 1, "Age").Return(nil)

				// Data writing expectations
				mock.EXPECT().ProcessValue("John", "").Return("John", nil)
				mock.EXPECT().SetCellValue(1, 2, "John").Return(nil)
				mock.EXPECT().ProcessValue(30, "").Return(30, nil)
				mock.EXPECT().SetCellValue(2, 2, 30).Return(nil)

				mock.EXPECT().ProcessValue("Jane", "").Return("Jane", nil)
				mock.EXPECT().SetCellValue(1, 3, "Jane").Return(nil)
				mock.EXPECT().ProcessValue(25, "").Return(25, nil)
				mock.EXPECT().SetCellValue(2, 3, 25).Return(nil)

				// Auto-fit columns expectations
				mock.EXPECT().GetColumnLetter(1).Return("A")
				mock.EXPECT().SetColumnWidth("A", 15.0).Return(nil)
				mock.EXPECT().GetColumnLetter(2).Return("B")
				mock.EXPECT().SetColumnWidth("B", 15.0).Return(nil)

				// Table processing expectations
				mock.EXPECT().ApplyBorderToCell(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().ApplyBordersToRange(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().ApplyStyleToCell(gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().ApplyStyleToRange(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().MergeCells(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().IsCellMerged(gomock.Any(), gomock.Any()).Return(false).AnyTimes()
				mock.EXPECT().IsCellMergedHorizontally(gomock.Any(), gomock.Any()).Return(false).AnyTimes()
				mock.EXPECT().HasExistingBorder(gomock.Any(), gomock.Any(), gomock.Any()).Return(false).AnyTimes()
				mock.EXPECT().GetCellValue(gomock.Any(), gomock.Any()).Return("", nil).AnyTimes()
			},
			expectError: false,
		},
		{
			name: "no_table_data",
			xlsx: &xlsx{
				params: FileWriteParams{},
			},
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().GetSheetName().Return("")
				mock.EXPECT().SetSheetName("Sheet1")
				mock.EXPECT().CreateSheet().Return(nil)
				mock.EXPECT().SetActiveSheet().Return(nil)
				mock.EXPECT().GetTable().Return(nil)
			},
			expectError: true,
			errorMsg:    "no table data provided",
		},
		{
			name: "create_sheet_error",
			xlsx: &xlsx{
				params: FileWriteParams{},
			},
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().GetSheetName().Return("")
				mock.EXPECT().SetSheetName("Sheet1")
				mock.EXPECT().CreateSheet().Return(errors.New("sheet creation error"))
			},
			expectError: true,
			errorMsg:    "failed to create sheet",
		},
		{
			name: "set_active_sheet_error",
			xlsx: &xlsx{
				params: FileWriteParams{},
			},
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().GetSheetName().Return("")
				mock.EXPECT().SetSheetName("Sheet1")
				mock.EXPECT().CreateSheet().Return(nil)
				mock.EXPECT().SetActiveSheet().Return(errors.New("set active sheet error"))
			},
			expectError: true,
			errorMsg:    "failed to set active sheet",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ctrl := gomock.NewController(t)
			defer ctrl.Finish()

			mockSpreadsheet := NewMockSpreadsheet(ctrl)
			tt.xlsx.spreadsheet = mockSpreadsheet
			tt.setupMock(mockSpreadsheet)

			err := tt.xlsx.writeData()

			if tt.expectError {
				if err == nil {
					t.Errorf("Expected error but got none")
				} else if tt.errorMsg != "" && !contains(err.Error(), tt.errorMsg) {
					t.Errorf("Expected error message to contain %q, got %q", tt.errorMsg, err.Error())
				}
			} else {
				if err != nil {
					t.Errorf("Unexpected error: %v", err)
				}
			}
		})
	}
}

// TestXlsx_writeHeaders tests the writeHeaders method
func TestXlsx_writeHeaders(t *testing.T) {
	tests := []struct {
		name         string
		xlsx         *xlsx
		setupMock    func(*MockSpreadsheet)
		expectError  bool
		errorMsg     string
		expectedRows int
	}{
		{
			name: "successful_simple_headers",
			xlsx: &xlsx{},
			setupMock: func(mock *MockSpreadsheet) {
				table := &Table{
					Columns: Columns{
						{Name: "name", Label: "Name"},
						{Name: "age", Label: "Age"},
					},
				}
				mock.EXPECT().GetTable().Return(table).AnyTimes()
				mock.EXPECT().SetCellValue(1, 1, "Name").Return(nil)
				mock.EXPECT().SetCellValue(2, 1, "Age").Return(nil)
			},
			expectError:  false,
			expectedRows: 1,
		},
		{
			name: "successful_multi_level_headers",
			xlsx: &xlsx{},
			setupMock: func(mock *MockSpreadsheet) {
				table := &Table{
					Columns: Columns{
						{
							Name:  "personal",
							Label: "Personal Info",
							Columns: Columns{
								{Name: "name", Label: "Name"},
								{Name: "age", Label: "Age"},
							},
						},
					},
				}
				mock.EXPECT().GetTable().Return(table).AnyTimes()
				mock.EXPECT().SetCellValue(1, 1, "Personal Info").Return(nil)
				mock.EXPECT().SetCellValue(1, 2, "Name").Return(nil)
				mock.EXPECT().SetCellValue(2, 2, "Age").Return(nil)
			},
			expectError:  false,
			expectedRows: 2,
		},
		{
			name: "no_columns",
			xlsx: &xlsx{},
			setupMock: func(mock *MockSpreadsheet) {
				table := &Table{
					Columns: Columns{},
				}
				mock.EXPECT().GetTable().Return(table).AnyTimes()
			},
			expectError:  false,
			expectedRows: 0,
		},
		{
			name: "header_cell_write_error",
			xlsx: &xlsx{},
			setupMock: func(mock *MockSpreadsheet) {
				table := &Table{
					Columns: Columns{
						{Name: "name", Label: "Name"},
					},
				}
				mock.EXPECT().GetTable().Return(table).AnyTimes()
				mock.EXPECT().SetCellValue(1, 1, "Name").Return(errors.New("cell write error"))
			},
			expectError: true,
			errorMsg:    "failed to set header cell value",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ctrl := gomock.NewController(t)
			defer ctrl.Finish()

			mockSpreadsheet := NewMockSpreadsheet(ctrl)
			tt.xlsx.spreadsheet = mockSpreadsheet
			tt.setupMock(mockSpreadsheet)

			rows, err := tt.xlsx.writeHeaders(1)

			if tt.expectError {
				if err == nil {
					t.Errorf("Expected error but got none")
				} else if tt.errorMsg != "" && !contains(err.Error(), tt.errorMsg) {
					t.Errorf("Expected error message to contain %q, got %q", tt.errorMsg, err.Error())
				}
			} else {
				if err != nil {
					t.Errorf("Unexpected error: %v", err)
				}
				if rows != tt.expectedRows {
					t.Errorf("Expected %d rows, got %d", tt.expectedRows, rows)
				}
			}
		})
	}
}

// TestXlsx_writeHeaderRow tests the writeHeaderRow method
func TestXlsx_writeHeaderRow(t *testing.T) {
	tests := []struct {
		name        string
		xlsx        *xlsx
		columns     Columns
		currentRow  int
		maxDepth    int
		startCol    int
		setupMock   func(*MockSpreadsheet)
		expectError bool
		errorMsg    string
	}{
		{
			name: "successful_simple_header_row",
			xlsx: &xlsx{},
			columns: Columns{
				{Name: "name", Label: "Name"},
				{Name: "age", Label: "Age"},
			},
			currentRow: 1,
			maxDepth:   1,
			startCol:   1,
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().SetCellValue(1, 1, "Name").Return(nil)
				mock.EXPECT().SetCellValue(2, 1, "Age").Return(nil)
			},
			expectError: false,
		},
		{
			name: "successful_hierarchical_header_row",
			xlsx: &xlsx{},
			columns: Columns{
				{
					Name:  "personal",
					Label: "Personal Info",
					Columns: Columns{
						{Name: "name", Label: "Name"},
						{Name: "age", Label: "Age"},
					},
				},
			},
			currentRow: 1,
			maxDepth:   2,
			startCol:   1,
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().SetCellValue(1, 1, "Personal Info").Return(nil)
				mock.EXPECT().SetCellValue(1, 2, "Name").Return(nil)
				mock.EXPECT().SetCellValue(2, 2, "Age").Return(nil)
			},
			expectError: false,
		},
		{
			name: "cell_write_error",
			xlsx: &xlsx{},
			columns: Columns{
				{Name: "name", Label: "Name"},
			},
			currentRow: 1,
			maxDepth:   1,
			startCol:   1,
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().SetCellValue(1, 1, "Name").Return(errors.New("cell write error"))
			},
			expectError: true,
			errorMsg:    "failed to set header cell value",
		},
		{
			name: "complex_nested_structure",
			xlsx: &xlsx{},
			columns: Columns{
				{
					Name:  "group1",
					Label: "Group 1",
					Columns: Columns{
						{Name: "col1", Label: "Column 1"},
					},
				},
				{Name: "col2", Label: "Column 2"},
			},
			currentRow: 1,
			maxDepth:   2,
			startCol:   1,
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().SetCellValue(1, 1, "Group 1").Return(nil)
				mock.EXPECT().SetCellValue(1, 2, "Column 1").Return(nil)
				mock.EXPECT().SetCellValue(2, 1, "Column 2").Return(nil)
			},
			expectError: false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ctrl := gomock.NewController(t)
			defer ctrl.Finish()

			mockSpreadsheet := NewMockSpreadsheet(ctrl)
			tt.xlsx.spreadsheet = mockSpreadsheet
			tt.setupMock(mockSpreadsheet)

			err := tt.xlsx.writeHeaderRow(tt.columns, tt.currentRow, tt.maxDepth, tt.startCol)

			if tt.expectError {
				if err == nil {
					t.Errorf("Expected error but got none")
				} else if tt.errorMsg != "" && !contains(err.Error(), tt.errorMsg) {
					t.Errorf("Expected error message to contain %q, got %q", tt.errorMsg, err.Error())
				}
			} else {
				if err != nil {
					t.Errorf("Unexpected error: %v", err)
				}
			}
		})
	}
}

// TestXlsx_writeCell tests the writeCell method
func TestXlsx_writeCell(t *testing.T) {
	tests := []struct {
		name        string
		xlsx        *xlsx
		item        Data
		column      *Column
		colIndex    int
		rowIndex    int
		setupMock   func(*MockSpreadsheet)
		expectError bool
		errorMsg    string
	}{
		{
			name: "successful_cell_write",
			xlsx: &xlsx{},
			item: Data{"name": "John", "age": 30},
			column: &Column{
				Name:   "name",
				Label:  "Name",
				Format: "",
			},
			colIndex: 1,
			rowIndex: 1,
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().ProcessValue("John", "").Return("John", nil)
				mock.EXPECT().SetCellValue(1, 1, "John").Return(nil)
			},
			expectError: false,
		},
		{
			name: "successful_cell_write_with_format",
			xlsx: &xlsx{},
			item: Data{"price": 123.45},
			column: &Column{
				Name:   "price",
				Label:  "Price",
				Format: "currency",
			},
			colIndex: 1,
			rowIndex: 1,
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().ProcessValue(123.45, "currency").Return("$123.45", nil)
				mock.EXPECT().SetCellValue(1, 1, "$123.45").Return(nil)
			},
			expectError: false,
		},
		{
			name: "missing_column_data",
			xlsx: &xlsx{},
			item: Data{"age": 30},
			column: &Column{
				Name:   "name",
				Label:  "Name",
				Format: "",
			},
			colIndex: 1,
			rowIndex: 1,
			setupMock: func(mock *MockSpreadsheet) {
				// No expectations as Lookup will skip
			},
			expectError: false,
		},
		{
			name: "process_value_error",
			xlsx: &xlsx{},
			item: Data{"name": "John"},
			column: &Column{
				Name:   "name",
				Label:  "Name",
				Format: "invalid_format",
			},
			colIndex: 1,
			rowIndex: 1,
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().ProcessValue("John", "invalid_format").Return(nil, errors.New("format error"))
			},
			expectError: true,
			errorMsg:    "error processing value",
		},
		{
			name: "set_cell_value_error",
			xlsx: &xlsx{},
			item: Data{"name": "John"},
			column: &Column{
				Name:   "name",
				Label:  "Name",
				Format: "",
			},
			colIndex: 1,
			rowIndex: 1,
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().ProcessValue("John", "").Return("John", nil)
				mock.EXPECT().SetCellValue(1, 1, "John").Return(errors.New("cell write error"))
			},
			expectError: true,
			errorMsg:    "error setting cell value",
		},
		{
			name: "nil_value_handling",
			xlsx: &xlsx{},
			item: Data{"name": nil},
			column: &Column{
				Name:   "name",
				Label:  "Name",
				Format: "",
			},
			colIndex: 1,
			rowIndex: 1,
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().ProcessValue(nil, "").Return("", nil)
				mock.EXPECT().SetCellValue(1, 1, "").Return(nil)
			},
			expectError: false,
		},
		{
			name: "formula_format_uses_set_cell_formula",
			xlsx: &xlsx{},
			item: Data{"total": "=SUM(A2:A10)"},
			column: &Column{
				Name:   "total",
				Label:  "Total",
				Format: ExcelizeFormatFormula,
			},
			colIndex: 1,
			rowIndex: 2,
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().ProcessValue("=SUM(A2:A10)", ExcelizeFormatFormula).Return("=SUM(A2:A10)", nil)
				mock.EXPECT().SetCellFormula(1, 2, "=SUM(A2:A10)").Return(nil)
			},
			expectError: false,
		},
		{
			name: "formula_format_set_cell_formula_error",
			xlsx: &xlsx{},
			item: Data{"total": "=SUM(A2:A10)"},
			column: &Column{
				Name:   "total",
				Label:  "Total",
				Format: ExcelizeFormatFormula,
			},
			colIndex: 1,
			rowIndex: 2,
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().ProcessValue("=SUM(A2:A10)", ExcelizeFormatFormula).Return("=SUM(A2:A10)", nil)
				mock.EXPECT().SetCellFormula(1, 2, "=SUM(A2:A10)").Return(errors.New("formula error"))
			},
			expectError: true,
			errorMsg:    "error setting formula",
		},
		{
			name: "hyperlink_format_sets_value_and_hyperlink",
			xlsx: &xlsx{},
			item: Data{"url": "https://example.com"},
			column: &Column{
				Name:   "url",
				Label:  "URL",
				Format: ExcelizeFormatHyperlink,
			},
			colIndex: 2,
			rowIndex: 3,
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().ProcessValue("https://example.com", ExcelizeFormatHyperlink).Return("https://example.com", nil)
				mock.EXPECT().SetCellValue(2, 3, "https://example.com").Return(nil)
				mock.EXPECT().SetCellHyperLink(2, 3, "https://example.com").Return(nil)
			},
			expectError: false,
		},
		{
			name: "hyperlink_format_set_cell_value_error",
			xlsx: &xlsx{},
			item: Data{"url": "https://example.com"},
			column: &Column{
				Name:   "url",
				Label:  "URL",
				Format: ExcelizeFormatHyperlink,
			},
			colIndex: 2,
			rowIndex: 3,
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().ProcessValue("https://example.com", ExcelizeFormatHyperlink).Return("https://example.com", nil)
				mock.EXPECT().SetCellValue(2, 3, "https://example.com").Return(errors.New("value error"))
			},
			expectError: true,
			errorMsg:    "error setting cell value",
		},
		{
			name: "hyperlink_format_set_cell_hyperlink_error",
			xlsx: &xlsx{},
			item: Data{"url": "https://example.com"},
			column: &Column{
				Name:   "url",
				Label:  "URL",
				Format: ExcelizeFormatHyperlink,
			},
			colIndex: 2,
			rowIndex: 3,
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().ProcessValue("https://example.com", ExcelizeFormatHyperlink).Return("https://example.com", nil)
				mock.EXPECT().SetCellValue(2, 3, "https://example.com").Return(nil)
				mock.EXPECT().SetCellHyperLink(2, 3, "https://example.com").Return(errors.New("hyperlink error"))
			},
			expectError: true,
			errorMsg:    "error setting hyperlink",
		},
		{
			name: "default_format_passes_raw_value",
			xlsx: &xlsx{},
			item: Data{"count": 42},
			column: &Column{
				Name:   "count",
				Label:  "Count",
				Format: ExcelizeFormatDefault,
			},
			colIndex: 1,
			rowIndex: 1,
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().ProcessValue(42, ExcelizeFormatDefault).Return(42, nil)
				mock.EXPECT().SetCellValue(1, 1, 42).Return(nil)
			},
			expectError: false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ctrl := gomock.NewController(t)
			defer ctrl.Finish()

			mockSpreadsheet := NewMockSpreadsheet(ctrl)
			tt.xlsx.spreadsheet = mockSpreadsheet
			tt.setupMock(mockSpreadsheet)

			err := tt.xlsx.writeCell(tt.item, tt.column, tt.colIndex, tt.rowIndex)

			if tt.expectError {
				if err == nil {
					t.Errorf("Expected error but got none")
				} else if tt.errorMsg != "" && !contains(err.Error(), tt.errorMsg) {
					t.Errorf("Expected error message to contain %q, got %q", tt.errorMsg, err.Error())
				}
			} else {
				if err != nil {
					t.Errorf("Unexpected error: %v", err)
				}
			}
		})
	}
}

// TestXlsx_autoFitColumns tests the autoFitColumns method
func TestXlsx_autoFitColumns(t *testing.T) {
	tests := []struct {
		name      string
		xlsx      *xlsx
		setupMock func(*MockSpreadsheet)
	}{
		{
			name: "successful_auto_fit",
			xlsx: &xlsx{},
			setupMock: func(mock *MockSpreadsheet) {
				table := &Table{
					Columns: Columns{
						{Name: "name", Label: "Name"},
						{Name: "age", Label: "Age"},
					},
				}
				mock.EXPECT().GetTable().Return(table).AnyTimes()
				mock.EXPECT().GetColumnLetter(1).Return("A")
				mock.EXPECT().SetColumnWidth("A", 15.0).Return(nil)
				mock.EXPECT().GetColumnLetter(2).Return("B")
				mock.EXPECT().SetColumnWidth("B", 15.0).Return(nil)
			},
		},
		{
			name: "column_width_error_continues",
			xlsx: &xlsx{},
			setupMock: func(mock *MockSpreadsheet) {
				table := &Table{
					Columns: Columns{
						{Name: "name", Label: "Name"},
						{Name: "age", Label: "Age"},
					},
				}
				mock.EXPECT().GetTable().Return(table).AnyTimes()
				mock.EXPECT().GetColumnLetter(1).Return("A")
				mock.EXPECT().SetColumnWidth("A", 15.0).Return(errors.New("width error"))
				mock.EXPECT().GetColumnLetter(2).Return("B")
				mock.EXPECT().SetColumnWidth("B", 15.0).Return(nil)
			},
		},
		{
			name: "no_columns",
			xlsx: &xlsx{},
			setupMock: func(mock *MockSpreadsheet) {
				table := &Table{
					Columns: Columns{},
				}
				mock.EXPECT().GetTable().Return(table).AnyTimes()
			},
		},
		{
			name: "single_column",
			xlsx: &xlsx{},
			setupMock: func(mock *MockSpreadsheet) {
				table := &Table{
					Columns: Columns{
						{Name: "single", Label: "Single Column"},
					},
				}
				mock.EXPECT().GetTable().Return(table).AnyTimes()
				mock.EXPECT().GetColumnLetter(1).Return("A")
				mock.EXPECT().SetColumnWidth("A", 15.0).Return(nil)
			},
		},
		{
			name: "many_columns",
			xlsx: &xlsx{},
			setupMock: func(mock *MockSpreadsheet) {
				table := &Table{
					Columns: Columns{
						{Name: "col1", Label: "Column 1"},
						{Name: "col2", Label: "Column 2"},
						{Name: "col3", Label: "Column 3"},
						{Name: "col4", Label: "Column 4"},
						{Name: "col5", Label: "Column 5"},
					},
				}
				mock.EXPECT().GetTable().Return(table).AnyTimes()
				mock.EXPECT().GetColumnLetter(1).Return("A")
				mock.EXPECT().SetColumnWidth("A", 15.0).Return(nil)
				mock.EXPECT().GetColumnLetter(2).Return("B")
				mock.EXPECT().SetColumnWidth("B", 15.0).Return(nil)
				mock.EXPECT().GetColumnLetter(3).Return("C")
				mock.EXPECT().SetColumnWidth("C", 15.0).Return(nil)
				mock.EXPECT().GetColumnLetter(4).Return("D")
				mock.EXPECT().SetColumnWidth("D", 15.0).Return(nil)
				mock.EXPECT().GetColumnLetter(5).Return("E")
				mock.EXPECT().SetColumnWidth("E", 15.0).Return(nil)
			},
		},
		{
			name: "custom_column_width",
			xlsx: &xlsx{},
			setupMock: func(mock *MockSpreadsheet) {
				table := &Table{
					Columns: Columns{
						{Name: "short", Label: "Short", Width: 0},  // 0 → default 15
						{Name: "long", Label: "Long Description", Width: 40},
					},
				}
				mock.EXPECT().GetTable().Return(table).AnyTimes()
				mock.EXPECT().GetColumnLetter(1).Return("A")
				mock.EXPECT().SetColumnWidth("A", 15.0).Return(nil)
				mock.EXPECT().GetColumnLetter(2).Return("B")
				mock.EXPECT().SetColumnWidth("B", 40.0).Return(nil)
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ctrl := gomock.NewController(t)
			defer ctrl.Finish()

			mockSpreadsheet := NewMockSpreadsheet(ctrl)
			tt.xlsx.spreadsheet = mockSpreadsheet
			tt.setupMock(mockSpreadsheet)

			// autoFitColumns doesn't return an error, so we just call it
			tt.xlsx.autoFitColumns()
		})
	}
}

// TestXlsx_writePreamble tests the writePreamble method
func TestXlsx_writePreamble(t *testing.T) {
	tests := []struct {
		name         string
		xlsx         *xlsx
		startRow     int
		setupMock    func(*MockSpreadsheet)
		expectError  bool
		errorMsg     string
		expectedRows int
	}{
		{
			name:     "no_preamble_rows",
			xlsx:     &xlsx{},
			startRow: 1,
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().GetTable().Return(&Table{}).AnyTimes()
			},
			expectedRows: 0,
		},
		{
			name:     "single_preamble_row",
			xlsx:     &xlsx{},
			startRow: 1,
			setupMock: func(mock *MockSpreadsheet) {
				table := &Table{
					Preamble: PreambleRows{
						NewPreambleRow("Report Title"),
					},
				}
				mock.EXPECT().GetTable().Return(table).AnyTimes()
				mock.EXPECT().SetCellValue(1, 1, "Report Title").Return(nil)
			},
			expectedRows: 1,
		},
		{
			name:     "multiple_preamble_rows_with_multiple_values",
			xlsx:     &xlsx{},
			startRow: 1,
			setupMock: func(mock *MockSpreadsheet) {
				table := &Table{
					Preamble: PreambleRows{
						NewPreambleRow("Title", "Date"),
						NewPreambleRow("Subtitle"),
					},
				}
				mock.EXPECT().GetTable().Return(table).AnyTimes()
				mock.EXPECT().SetCellValue(1, 1, "Title").Return(nil)
				mock.EXPECT().SetCellValue(2, 1, "Date").Return(nil)
				mock.EXPECT().SetCellValue(1, 2, "Subtitle").Return(nil)
			},
			expectedRows: 2,
		},
		{
			name:     "preamble_with_offset_start_row",
			xlsx:     &xlsx{},
			startRow: 3,
			setupMock: func(mock *MockSpreadsheet) {
				table := &Table{
					Preamble: PreambleRows{
						NewPreambleRow("Title"),
					},
				}
				mock.EXPECT().GetTable().Return(table).AnyTimes()
				mock.EXPECT().SetCellValue(1, 3, "Title").Return(nil)
			},
			expectedRows: 1,
		},
		{
			name:     "preamble_cell_write_error",
			xlsx:     &xlsx{},
			startRow: 1,
			setupMock: func(mock *MockSpreadsheet) {
				table := &Table{
					Preamble: PreambleRows{
						NewPreambleRow("Title"),
					},
				}
				mock.EXPECT().GetTable().Return(table).AnyTimes()
				mock.EXPECT().SetCellValue(1, 1, "Title").Return(errors.New("write error"))
			},
			expectError: true,
			errorMsg:    "failed to write preamble cell",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ctrl := gomock.NewController(t)
			defer ctrl.Finish()

			mockSpreadsheet := NewMockSpreadsheet(ctrl)
			tt.xlsx.spreadsheet = mockSpreadsheet
			tt.setupMock(mockSpreadsheet)

			rows, err := tt.xlsx.writePreamble(tt.startRow)

			if tt.expectError {
				if err == nil {
					t.Errorf("Expected error but got none")
				} else if tt.errorMsg != "" && !contains(err.Error(), tt.errorMsg) {
					t.Errorf("Expected error message to contain %q, got %q", tt.errorMsg, err.Error())
				}
			} else {
				if err != nil {
					t.Errorf("Unexpected error: %v", err)
				}
				if rows != tt.expectedRows {
					t.Errorf("Expected %d rows, got %d", tt.expectedRows, rows)
				}
			}
		})
	}
}

// TestXlsx_writeHeaders_withPreamble tests writeHeaders when preamble rows push headers down
func TestXlsx_writeHeaders_withPreamble(t *testing.T) {
	tests := []struct {
		name         string
		xlsx         *xlsx
		startRow     int
		setupMock    func(*MockSpreadsheet)
		expectError  bool
		expectedRows int
	}{
		{
			name:     "headers_start_at_row_2_due_to_preamble",
			xlsx:     &xlsx{},
			startRow: 2,
			setupMock: func(mock *MockSpreadsheet) {
				table := &Table{
					Columns: Columns{
						{Name: "name", Label: "Name"},
						{Name: "age", Label: "Age"},
					},
				}
				mock.EXPECT().GetTable().Return(table).AnyTimes()
				mock.EXPECT().SetCellValue(1, 2, "Name").Return(nil)
				mock.EXPECT().SetCellValue(2, 2, "Age").Return(nil)
			},
			expectedRows: 1,
		},
		{
			name:     "multi_level_headers_with_preamble",
			xlsx:     &xlsx{},
			startRow: 3,
			setupMock: func(mock *MockSpreadsheet) {
				table := &Table{
					Columns: Columns{
						{
							Name:  "personal",
							Label: "Personal",
							Columns: Columns{
								{Name: "name", Label: "Name"},
							},
						},
					},
				}
				mock.EXPECT().GetTable().Return(table).AnyTimes()
				mock.EXPECT().SetCellValue(1, 3, "Personal").Return(nil)
				mock.EXPECT().SetCellValue(1, 4, "Name").Return(nil)
			},
			expectedRows: 2,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ctrl := gomock.NewController(t)
			defer ctrl.Finish()

			mockSpreadsheet := NewMockSpreadsheet(ctrl)
			tt.xlsx.spreadsheet = mockSpreadsheet
			tt.setupMock(mockSpreadsheet)

			rows, err := tt.xlsx.writeHeaders(tt.startRow)

			if tt.expectError {
				if err == nil {
					t.Errorf("Expected error but got none")
				}
			} else {
				if err != nil {
					t.Errorf("Unexpected error: %v", err)
				}
				if rows != tt.expectedRows {
					t.Errorf("Expected %d rows, got %d", tt.expectedRows, rows)
				}
			}
		})
	}
}

// TestExportXLSXSheets tests the ExportXLSXSheets function
func TestExportXLSXSheets(t *testing.T) {
	ctrl := gomock.NewController(t)
	defer ctrl.Finish()

	tempDir, err := os.MkdirTemp("", "xlsx_sheets_test")
	if err != nil {
		t.Fatalf("Failed to create temp directory: %v", err)
	}
	defer func() {
		if err := os.RemoveAll(tempDir); err != nil {
			t.Logf("Failed to remove temp directory: %v", err)
		}
	}()

	emptyTable := &Table{Data: DataSlice{}, Columns: Columns{}, WriteHeader: false}

	tests := []struct {
		name          string
		setupMocks    func(mocks []*MockSpreadsheet)
		sheetCount    int
		params        FileWriteParams
		expectError   bool
		errorContains string
	}{
		{
			name:          "no_sheets_returns_error",
			setupMocks:    func(mocks []*MockSpreadsheet) {},
			sheetCount:    0,
			params:        FileWriteParams{Filename: "test_no_sheets", Filepath: tempDir},
			expectError:   true,
			errorContains: "no sheets provided",
		},
		{
			name: "single_sheet_with_new_file",
			setupMocks: func(mocks []*MockSpreadsheet) {
				m := mocks[0]
				m.EXPECT().GetFile().Return(nil)
				m.EXPECT().CreateNewFile().Return(nil)
				m.EXPECT().Close().Return(nil)
				m.EXPECT().GetSheetName().Return("Reports")
				m.EXPECT().CreateSheet().Return(nil)
				m.EXPECT().SetActiveSheet().Return(nil)
				m.EXPECT().GetTable().Return(emptyTable).AnyTimes()
				m.EXPECT().SaveToWriter(gomock.Any()).Return(nil)
			},
			sheetCount: 1,
			params:     FileWriteParams{Filename: "test_single_sheet", Filepath: tempDir},
		},
		{
			name: "multiple_sheets_with_new_file",
			setupMocks: func(mocks []*MockSpreadsheet) {
				existingFile := &struct{ name string }{name: "shared_file"}
				mocks[0].EXPECT().GetFile().Return(nil)
				mocks[0].EXPECT().CreateNewFile().Return(nil)
				mocks[0].EXPECT().Close().Return(nil)
				mocks[0].EXPECT().GetFile().Return(existingFile) // second call to propagate file
				mocks[0].EXPECT().GetSheetName().Return("Sheet1")
				mocks[0].EXPECT().CreateSheet().Return(nil)
				mocks[0].EXPECT().SetActiveSheet().Return(nil)
				mocks[0].EXPECT().GetTable().Return(emptyTable).AnyTimes()
				mocks[0].EXPECT().SaveToWriter(gomock.Any()).Return(nil)

				mocks[1].EXPECT().GetFile().Return(nil)
				mocks[1].EXPECT().InitWithFile(existingFile).Return(nil)
				mocks[1].EXPECT().GetSheetName().Return("Summary")
				mocks[1].EXPECT().CreateSheet().Return(nil)
				mocks[1].EXPECT().SetActiveSheet().Return(nil)
				mocks[1].EXPECT().GetTable().Return(emptyTable).AnyTimes()
			},
			sheetCount: 2,
			params:     FileWriteParams{Filename: "test_multiple_sheets", Filepath: tempDir},
		},
		{
			name: "multiple_sheets_with_existing_file",
			setupMocks: func(mocks []*MockSpreadsheet) {
				existingFile := &struct{ name string }{name: "shared_file"}
				mocks[0].EXPECT().GetFile().Return(existingFile)
				mocks[0].EXPECT().GetFile().Return(existingFile) // second call to propagate file
				mocks[0].EXPECT().GetSheetName().Return("Sheet1")
				mocks[0].EXPECT().CreateSheet().Return(nil)
				mocks[0].EXPECT().SetActiveSheet().Return(nil)
				mocks[0].EXPECT().GetTable().Return(emptyTable).AnyTimes()
				mocks[0].EXPECT().SaveToWriter(gomock.Any()).Return(nil)

				mocks[1].EXPECT().GetFile().Return(existingFile) // already has the file
				mocks[1].EXPECT().GetSheetName().Return("Summary")
				mocks[1].EXPECT().CreateSheet().Return(nil)
				mocks[1].EXPECT().SetActiveSheet().Return(nil)
				mocks[1].EXPECT().GetTable().Return(emptyTable).AnyTimes()
			},
			sheetCount: 2,
			params:     FileWriteParams{Filename: "test_existing_file", Filepath: tempDir},
		},
		{
			name: "error_creating_new_file",
			setupMocks: func(mocks []*MockSpreadsheet) {
				mocks[0].EXPECT().GetFile().Return(nil)
				mocks[0].EXPECT().CreateNewFile().Return(errors.New("create file error"))
			},
			sheetCount:    1,
			params:        FileWriteParams{Filename: "test_error_create", Filepath: tempDir},
			expectError:   true,
			errorContains: "failed to create new XLSX file",
		},
		{
			name: "error_initializing_second_sheet",
			setupMocks: func(mocks []*MockSpreadsheet) {
				existingFile := &struct{ name string }{name: "shared_file"}
				mocks[0].EXPECT().GetFile().Return(nil)
				mocks[0].EXPECT().CreateNewFile().Return(nil)
				mocks[0].EXPECT().Close().Return(nil)
				mocks[0].EXPECT().GetFile().Return(existingFile) // second call to propagate file

				mocks[1].EXPECT().GetFile().Return(nil)
				mocks[1].EXPECT().InitWithFile(existingFile).Return(errors.New("init error"))
			},
			sheetCount:    2,
			params:        FileWriteParams{Filename: "test_error_init", Filepath: tempDir},
			expectError:   true,
			errorContains: "failed to initialize sheet with existing file",
		},
		{
			name: "error_writing_data",
			setupMocks: func(mocks []*MockSpreadsheet) {
				mocks[0].EXPECT().GetFile().Return(&struct{ name string }{name: "existing_file"})
				mocks[0].EXPECT().GetSheetName().Return("Sheet1")
				mocks[0].EXPECT().CreateSheet().Return(errors.New("create sheet error"))
			},
			sheetCount:    1,
			params:        FileWriteParams{Filename: "test_write_error", Filepath: tempDir},
			expectError:   true,
			errorContains: "failed to write data to XLSX file",
		},
		{
			name: "error_saving_to_writer",
			setupMocks: func(mocks []*MockSpreadsheet) {
				mocks[0].EXPECT().GetFile().Return(&struct{ name string }{name: "existing_file"})
				mocks[0].EXPECT().GetSheetName().Return("Sheet1")
				mocks[0].EXPECT().CreateSheet().Return(nil)
				mocks[0].EXPECT().SetActiveSheet().Return(nil)
				mocks[0].EXPECT().GetTable().Return(emptyTable).AnyTimes()
				mocks[0].EXPECT().SaveToWriter(gomock.Any()).Return(errors.New("save error"))
			},
			sheetCount:    1,
			params:        FileWriteParams{Filename: "test_save_error", Filepath: tempDir},
			expectError:   true,
			errorContains: "failed to write XLSX to writer",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			mocks := make([]*MockSpreadsheet, tt.sheetCount)
			for i := range mocks {
				mocks[i] = NewMockSpreadsheet(ctrl)
			}
			tt.setupMocks(mocks)

			sheets := make([]Spreadsheet, tt.sheetCount)
			for i, m := range mocks {
				sheets[i] = m
			}

			result, err := ExportXLSXSheets(sheets, tt.params)

			if tt.expectError {
				if err == nil {
					t.Error("Expected error but got none")
				} else if tt.errorContains != "" && !contains(err.Error(), tt.errorContains) {
					t.Errorf("Expected error to contain %q, got %q", tt.errorContains, err.Error())
				}
			} else {
				if err != nil {
					t.Errorf("Unexpected error: %v", err)
				}
				if result == nil {
					t.Error("Expected non-nil result")
				}
			}
		})
	}
}

// Helper function to check if a string contains a substring
func contains(s, substr string) bool {
	return len(s) >= len(substr) && (s == substr || len(substr) == 0 ||
		(len(s) > len(substr) &&
			(s[:len(substr)] == substr || s[len(s)-len(substr):] == substr ||
				func() bool {
					for i := 0; i <= len(s)-len(substr); i++ {
						if s[i:i+len(substr)] == substr {
							return true
						}
					}
					return false
				}())))
}
