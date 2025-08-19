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
				mock.EXPECT().getFile().Return(nil)
				mock.EXPECT().createNewFile().Return(nil)
				mock.EXPECT().close().Return(nil)
				mock.EXPECT().getSheetName().Return("")
				mock.EXPECT().setSheetName("Sheet1")
				mock.EXPECT().createSheet().Return(nil)
				mock.EXPECT().setActiveSheet().Return(nil)
				mock.EXPECT().getTable().Return(&Table{
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
				mock.EXPECT().setCellValue(1, 1, "Name").Return(nil)
				mock.EXPECT().setCellValue(2, 1, "Age").Return(nil)

				// Data writing expectations
				mock.EXPECT().processValue("John", "").Return("John", nil)
				mock.EXPECT().setCellValue(1, 2, "John").Return(nil)
				mock.EXPECT().processValue(30, "").Return(30, nil)
				mock.EXPECT().setCellValue(2, 2, 30).Return(nil)

				// Auto-fit columns expectations
				mock.EXPECT().getColumnLetter(1).Return("A")
				mock.EXPECT().setColumnWidth("A", 15.0).Return(nil)
				mock.EXPECT().getColumnLetter(2).Return("B")
				mock.EXPECT().setColumnWidth("B", 15.0).Return(nil)

				// Table processing expectations (merging and styling are called on the table, not the spreadsheet)
				// We need to mock any calls that these methods might make to the spreadsheet
				mock.EXPECT().applyBorderToCell(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().applyBordersToRange(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().applyStyleToCell(gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().applyStyleToRange(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().mergeCells(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().isCellMerged(gomock.Any(), gomock.Any()).Return(false).AnyTimes()
				mock.EXPECT().isCellMergedHorizontally(gomock.Any(), gomock.Any()).Return(false).AnyTimes()
				mock.EXPECT().hasExistingBorder(gomock.Any(), gomock.Any(), gomock.Any()).Return(false).AnyTimes()
				mock.EXPECT().getCellValue(gomock.Any(), gomock.Any()).Return("", nil).AnyTimes()
				mock.EXPECT().setCellValue(gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().getColumnLetter(gomock.Any()).Return("A").AnyTimes()
				mock.EXPECT().processValue(gomock.Any(), gomock.Any()).Return(gomock.Any(), nil).AnyTimes()

				// Final save expectation
				mock.EXPECT().saveToWriter(gomock.Any()).Return(nil)
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
				mock.EXPECT().getFile().Return(&struct{ name string }{name: "existing_file"})
				mock.EXPECT().getSheetName().Return("Sheet1")
				mock.EXPECT().createSheet().Return(nil)
				mock.EXPECT().setActiveSheet().Return(nil)
				mock.EXPECT().getTable().Return(&Table{
					Data:        DataSlice{},
					Columns:     Columns{},
					WriteHeader: false,
				}).AnyTimes()
				mock.EXPECT().saveToWriter(gomock.Any()).Return(nil)
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
				mock.EXPECT().getFile().Return(nil)
				mock.EXPECT().createNewFile().Return(errors.New("create file error"))
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
				mock.EXPECT().getFile().Return(&struct{ name string }{name: "existing_file"})
				mock.EXPECT().getSheetName().Return("Sheet1")
				mock.EXPECT().createSheet().Return(errors.New("create sheet error"))
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
				mock.EXPECT().getFile().Return(&struct{ name string }{name: "existing_file"})
				mock.EXPECT().getSheetName().Return("Sheet1")
				mock.EXPECT().createSheet().Return(nil)
				mock.EXPECT().setActiveSheet().Return(nil)
				mock.EXPECT().getTable().Return(&Table{
					Data:        DataSlice{},
					Columns:     Columns{},
					WriteHeader: false,
				}).AnyTimes()
				mock.EXPECT().saveToWriter(gomock.Any()).Return(errors.New("save error"))
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

				mock.EXPECT().getSheetName().Return("")
				mock.EXPECT().setSheetName("Sheet1")
				mock.EXPECT().createSheet().Return(nil)
				mock.EXPECT().setActiveSheet().Return(nil)
				mock.EXPECT().getTable().Return(table).AnyTimes()

				// Header writing expectations
				mock.EXPECT().setCellValue(1, 1, "Name").Return(nil)
				mock.EXPECT().setCellValue(2, 1, "Age").Return(nil)

				// Data writing expectations
				mock.EXPECT().processValue("John", "").Return("John", nil)
				mock.EXPECT().setCellValue(1, 2, "John").Return(nil)
				mock.EXPECT().processValue(30, "").Return(30, nil)
				mock.EXPECT().setCellValue(2, 2, 30).Return(nil)

				mock.EXPECT().processValue("Jane", "").Return("Jane", nil)
				mock.EXPECT().setCellValue(1, 3, "Jane").Return(nil)
				mock.EXPECT().processValue(25, "").Return(25, nil)
				mock.EXPECT().setCellValue(2, 3, 25).Return(nil)

				// Auto-fit columns expectations
				mock.EXPECT().getColumnLetter(1).Return("A")
				mock.EXPECT().setColumnWidth("A", 15.0).Return(nil)
				mock.EXPECT().getColumnLetter(2).Return("B")
				mock.EXPECT().setColumnWidth("B", 15.0).Return(nil)

				// Table processing expectations
				mock.EXPECT().applyBorderToCell(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().applyBordersToRange(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().applyStyleToCell(gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().applyStyleToRange(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().mergeCells(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().isCellMerged(gomock.Any(), gomock.Any()).Return(false).AnyTimes()
				mock.EXPECT().isCellMergedHorizontally(gomock.Any(), gomock.Any()).Return(false).AnyTimes()
				mock.EXPECT().hasExistingBorder(gomock.Any(), gomock.Any(), gomock.Any()).Return(false).AnyTimes()
				mock.EXPECT().getCellValue(gomock.Any(), gomock.Any()).Return("", nil).AnyTimes()
			},
			expectError: false,
		},
		{
			name: "no_table_data",
			xlsx: &xlsx{
				params: FileWriteParams{},
			},
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().getSheetName().Return("")
				mock.EXPECT().setSheetName("Sheet1")
				mock.EXPECT().createSheet().Return(nil)
				mock.EXPECT().setActiveSheet().Return(nil)
				mock.EXPECT().getTable().Return(nil)
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
				mock.EXPECT().getSheetName().Return("")
				mock.EXPECT().setSheetName("Sheet1")
				mock.EXPECT().createSheet().Return(errors.New("sheet creation error"))
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
				mock.EXPECT().getSheetName().Return("")
				mock.EXPECT().setSheetName("Sheet1")
				mock.EXPECT().createSheet().Return(nil)
				mock.EXPECT().setActiveSheet().Return(errors.New("set active sheet error"))
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
				mock.EXPECT().getTable().Return(table).AnyTimes()
				mock.EXPECT().setCellValue(1, 1, "Name").Return(nil)
				mock.EXPECT().setCellValue(2, 1, "Age").Return(nil)
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
				mock.EXPECT().getTable().Return(table).AnyTimes()
				mock.EXPECT().setCellValue(1, 1, "Personal Info").Return(nil)
				mock.EXPECT().setCellValue(1, 2, "Name").Return(nil)
				mock.EXPECT().setCellValue(2, 2, "Age").Return(nil)
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
				mock.EXPECT().getTable().Return(table).AnyTimes()
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
				mock.EXPECT().getTable().Return(table).AnyTimes()
				mock.EXPECT().setCellValue(1, 1, "Name").Return(errors.New("cell write error"))
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

			rows, err := tt.xlsx.writeHeaders()

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
				mock.EXPECT().setCellValue(1, 1, "Name").Return(nil)
				mock.EXPECT().setCellValue(2, 1, "Age").Return(nil)
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
				mock.EXPECT().setCellValue(1, 1, "Personal Info").Return(nil)
				mock.EXPECT().setCellValue(1, 2, "Name").Return(nil)
				mock.EXPECT().setCellValue(2, 2, "Age").Return(nil)
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
				mock.EXPECT().setCellValue(1, 1, "Name").Return(errors.New("cell write error"))
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
				mock.EXPECT().setCellValue(1, 1, "Group 1").Return(nil)
				mock.EXPECT().setCellValue(1, 2, "Column 1").Return(nil)
				mock.EXPECT().setCellValue(2, 1, "Column 2").Return(nil)
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
		column      Column
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
			column: Column{
				Name:   "name",
				Label:  "Name",
				Format: "",
			},
			colIndex: 1,
			rowIndex: 1,
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().processValue("John", "").Return("John", nil)
				mock.EXPECT().setCellValue(1, 1, "John").Return(nil)
			},
			expectError: false,
		},
		{
			name: "successful_cell_write_with_format",
			xlsx: &xlsx{},
			item: Data{"price": 123.45},
			column: Column{
				Name:   "price",
				Label:  "Price",
				Format: "currency",
			},
			colIndex: 1,
			rowIndex: 1,
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().processValue(123.45, "currency").Return("$123.45", nil)
				mock.EXPECT().setCellValue(1, 1, "$123.45").Return(nil)
			},
			expectError: false,
		},
		{
			name: "missing_column_data",
			xlsx: &xlsx{},
			item: Data{"age": 30},
			column: Column{
				Name:   "name",
				Label:  "Name",
				Format: "",
			},
			colIndex: 1,
			rowIndex: 1,
			setupMock: func(mock *MockSpreadsheet) {
				// No expectations as lookup will fail
			},
			expectError: true,
			errorMsg:    "error looking up value for column",
		},
		{
			name: "process_value_error",
			xlsx: &xlsx{},
			item: Data{"name": "John"},
			column: Column{
				Name:   "name",
				Label:  "Name",
				Format: "invalid_format",
			},
			colIndex: 1,
			rowIndex: 1,
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().processValue("John", "invalid_format").Return(nil, errors.New("format error"))
			},
			expectError: true,
			errorMsg:    "error processing value",
		},
		{
			name: "set_cell_value_error",
			xlsx: &xlsx{},
			item: Data{"name": "John"},
			column: Column{
				Name:   "name",
				Label:  "Name",
				Format: "",
			},
			colIndex: 1,
			rowIndex: 1,
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().processValue("John", "").Return("John", nil)
				mock.EXPECT().setCellValue(1, 1, "John").Return(errors.New("cell write error"))
			},
			expectError: true,
			errorMsg:    "error setting cell value",
		},
		{
			name: "nil_value_handling",
			xlsx: &xlsx{},
			item: Data{"name": nil},
			column: Column{
				Name:   "name",
				Label:  "Name",
				Format: "",
			},
			colIndex: 1,
			rowIndex: 1,
			setupMock: func(mock *MockSpreadsheet) {
				mock.EXPECT().processValue(nil, "").Return("", nil)
				mock.EXPECT().setCellValue(1, 1, "").Return(nil)
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
				mock.EXPECT().getTable().Return(table).AnyTimes()
				mock.EXPECT().getColumnLetter(1).Return("A")
				mock.EXPECT().setColumnWidth("A", 15.0).Return(nil)
				mock.EXPECT().getColumnLetter(2).Return("B")
				mock.EXPECT().setColumnWidth("B", 15.0).Return(nil)
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
				mock.EXPECT().getTable().Return(table).AnyTimes()
				mock.EXPECT().getColumnLetter(1).Return("A")
				mock.EXPECT().setColumnWidth("A", 15.0).Return(errors.New("width error"))
				mock.EXPECT().getColumnLetter(2).Return("B")
				mock.EXPECT().setColumnWidth("B", 15.0).Return(nil)
			},
		},
		{
			name: "no_columns",
			xlsx: &xlsx{},
			setupMock: func(mock *MockSpreadsheet) {
				table := &Table{
					Columns: Columns{},
				}
				mock.EXPECT().getTable().Return(table).AnyTimes()
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
				mock.EXPECT().getTable().Return(table).AnyTimes()
				mock.EXPECT().getColumnLetter(1).Return("A")
				mock.EXPECT().setColumnWidth("A", 15.0).Return(nil)
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
				mock.EXPECT().getTable().Return(table).AnyTimes()
				mock.EXPECT().getColumnLetter(1).Return("A")
				mock.EXPECT().setColumnWidth("A", 15.0).Return(nil)
				mock.EXPECT().getColumnLetter(2).Return("B")
				mock.EXPECT().setColumnWidth("B", 15.0).Return(nil)
				mock.EXPECT().getColumnLetter(3).Return("C")
				mock.EXPECT().setColumnWidth("C", 15.0).Return(nil)
				mock.EXPECT().getColumnLetter(4).Return("D")
				mock.EXPECT().setColumnWidth("D", 15.0).Return(nil)
				mock.EXPECT().getColumnLetter(5).Return("E")
				mock.EXPECT().setColumnWidth("E", 15.0).Return(nil)
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
