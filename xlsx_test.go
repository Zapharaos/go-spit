package spit_test

import (
	"errors"
	"os"
	"testing"

	"github.com/Zapharaos/go-spit"
	"github.com/Zapharaos/go-spit/mocks"
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
		setupMock      func(*mocks.MockSpreadsheet)
		params         spit.FileWriteParams
		expectError    bool
		errorContains  string
		validateResult func(*spit.FileWriteResult)
	}{
		{
			name: "successful_export_with_new_file",
			setupMock: func(mock *mocks.MockSpreadsheet) {
				mock.EXPECT().GetFile().Return(nil)
				mock.EXPECT().CreateNewFile().Return(nil)
				mock.EXPECT().Close().Return(nil)
				mock.EXPECT().GetSheetName().Return("")
				mock.EXPECT().SetSheetName("Sheet1")
				mock.EXPECT().CreateSheet().Return(nil)
				mock.EXPECT().SetActiveSheet().Return(nil)
				mock.EXPECT().GetTable().Return(&spit.Table{
					Data: spit.DataSlice{
						{"name": "John", "age": 30},
					},
					Columns: spit.Columns{
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
				mock.EXPECT().GetColumnLetter(2).Return("B")
				mock.EXPECT().SetColumnWidth("A", 15.0).Return(nil)
				mock.EXPECT().SetColumnWidth("B", 15.0).Return(nil)

				// Table processing expectations (merging and styling)
				// Allow any number of border/style calls since these depend on table configuration
				mock.EXPECT().ApplyBorderToCell(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().ApplyBordersToRange(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().ApplyStyleToCell(gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().ApplyStyleToRange(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().MergeCells(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().IsCellMerged(gomock.Any(), gomock.Any()).Return(false).AnyTimes()
				mock.EXPECT().IsCellMergedHorizontally(gomock.Any(), gomock.Any()).Return(false).AnyTimes()
				mock.EXPECT().HasExistingBorder(gomock.Any(), gomock.Any(), gomock.Any()).Return(false).AnyTimes()

				// Final save expectation
				mock.EXPECT().SaveToWriter(gomock.Any()).Return(nil)
			},
			params: spit.FileWriteParams{
				Filename: "test_export",
				Filepath: tempDir,
			},
			expectError: false,
			validateResult: func(result *spit.FileWriteResult) {
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
			setupMock: func(mock *mocks.MockSpreadsheet) {
				mock.EXPECT().GetFile().Return(&struct{ name string }{name: "existing_file"})
				mock.EXPECT().GetSheetName().Return("Sheet1")
				mock.EXPECT().CreateSheet().Return(nil)
				mock.EXPECT().SetActiveSheet().Return(nil)
				mock.EXPECT().GetTable().Return(&spit.Table{
					Data:        spit.DataSlice{},
					Columns:     spit.Columns{},
					WriteHeader: false,
				}).AnyTimes()
				mock.EXPECT().SaveToWriter(gomock.Any()).Return(nil)
			},
			params: spit.FileWriteParams{
				Filename: "test_existing",
				Filepath: tempDir,
			},
			expectError: false,
		},
		{
			name: "error_creating_new_file",
			setupMock: func(mock *mocks.MockSpreadsheet) {
				mock.EXPECT().GetFile().Return(nil)
				mock.EXPECT().CreateNewFile().Return(errors.New("create file error"))
			},
			params: spit.FileWriteParams{
				Filename: "test_error",
				Filepath: tempDir,
			},
			expectError:   true,
			errorContains: "failed to create new XLSX file",
		},
		{
			name: "error_writing_data",
			setupMock: func(mock *mocks.MockSpreadsheet) {
				mock.EXPECT().GetFile().Return(&struct{ name string }{name: "existing_file"})
				mock.EXPECT().GetSheetName().Return("Sheet1")
				mock.EXPECT().CreateSheet().Return(errors.New("create sheet error"))
			},
			params: spit.FileWriteParams{
				Filename: "test_write_error",
				Filepath: tempDir,
			},
			expectError:   true,
			errorContains: "failed to write data to XLSX file",
		},
		{
			name: "error_saving_to_writer",
			setupMock: func(mock *mocks.MockSpreadsheet) {
				mock.EXPECT().GetFile().Return(&struct{ name string }{name: "existing_file"})
				mock.EXPECT().GetSheetName().Return("Sheet1")
				mock.EXPECT().CreateSheet().Return(nil)
				mock.EXPECT().SetActiveSheet().Return(nil)
				mock.EXPECT().GetTable().Return(&spit.Table{
					Data:        spit.DataSlice{},
					Columns:     spit.Columns{},
					WriteHeader: false,
				}).AnyTimes()
				mock.EXPECT().SaveToWriter(gomock.Any()).Return(errors.New("save error"))
			},
			params: spit.FileWriteParams{
				Filename: "test_save_error",
				Filepath: tempDir,
			},
			expectError:   true,
			errorContains: "failed to write XLSX to writer",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			mockSpreadsheet := mocks.NewMockSpreadsheet(ctrl)
			tt.setupMock(mockSpreadsheet)

			result, err := spit.ExportXLSX(mockSpreadsheet, tt.params)

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
