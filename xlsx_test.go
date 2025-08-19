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
				mock.EXPECT().getColumnLetter(2).Return("B")
				mock.EXPECT().setColumnWidth("A", 15.0).Return(nil)
				mock.EXPECT().setColumnWidth("B", 15.0).Return(nil)

				// Table processing expectations (merging and styling)
				// Allow any number of border/style calls since these depend on table configuration
				mock.EXPECT().applyBorderToCell(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().applyBordersToRange(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().applyStyleToCell(gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().applyStyleToRange(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().mergeCells(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().isCellMerged(gomock.Any(), gomock.Any()).Return(false).AnyTimes()
				mock.EXPECT().isCellMergedHorizontally(gomock.Any(), gomock.Any()).Return(false).AnyTimes()
				mock.EXPECT().hasExistingBorder(gomock.Any(), gomock.Any(), gomock.Any()).Return(false).AnyTimes()

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
