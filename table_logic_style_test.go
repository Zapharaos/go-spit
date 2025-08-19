package spit

import (
	"errors"
	"testing"

	"go.uber.org/mock/gomock"
)

func TestTable_renderStyles(t *testing.T) {
	ctrl := gomock.NewController(t)
	defer ctrl.Finish()

	tests := []struct {
		name          string
		table         *Table
		setupMock     func(*MockTableOperations)
		expectedError bool
		errorContains string
	}{
		{
			name: "success_with_headers_and_data",
			table: &Table{
				Data: DataSlice{
					{"name": "John", "age": 30},
					{"name": "Jane", "age": 25},
				},
				Columns: Columns{
					{Name: "name", Label: "Name"},
					{Name: "age", Label: "Age"},
				},
				WriteHeader: true,
			},
			setupMock: func(mockOps *MockTableOperations) {
				// Header styles
				headerStyle := Style{
					Bold:            true,
					BackgroundColor: "#E0E0E0",
					Alignment:       AlignmentCenterMiddle,
				}
				mockOps.EXPECT().applyStyleToRange(1, 1, 2, 1, headerStyle).Return(nil)

				// Header borders (2 columns * 4 sides = 8 calls)
				border := &Border{Style: BorderStyleThin}
				for col := 1; col <= 2; col++ {
					for _, side := range []string{"left", "right", "top", "bottom"} {
						mockOps.EXPECT().applyBorderToCell(col, 1, side, border).Return(nil)
					}
				}
			},
			expectedError: false,
		},
		{
			name: "success_no_headers",
			table: &Table{
				Data: DataSlice{
					{"name": "John", "age": 30},
				},
				Columns: Columns{
					{Name: "name", Label: "Name"},
					{Name: "age", Label: "Age"},
				},
				WriteHeader: false,
			},
			setupMock: func(mockOps *MockTableOperations) {
				// No header style calls expected
			},
			expectedError: false,
		},
		{
			name: "error_in_header_styles",
			table: &Table{
				Data: DataSlice{
					{"name": "John"},
				},
				Columns: Columns{
					{Name: "name", Label: "Name"},
				},
				WriteHeader: true,
			},
			setupMock: func(mockOps *MockTableOperations) {
				// The header borders are applied first, then the style
				border := &Border{Style: BorderStyleThin}
				mockOps.EXPECT().applyBorderToCell(1, 1, "left", border).Return(nil)
				mockOps.EXPECT().applyBorderToCell(1, 1, "right", border).Return(nil)
				mockOps.EXPECT().applyBorderToCell(1, 1, "top", border).Return(nil)
				mockOps.EXPECT().applyBorderToCell(1, 1, "bottom", border).Return(nil)

				// Then the style fails
				headerStyle := Style{
					Bold:            true,
					BackgroundColor: "#E0E0E0",
					Alignment:       AlignmentCenterMiddle,
				}
				mockOps.EXPECT().applyStyleToRange(1, 1, 1, 1, headerStyle).Return(errors.New("header style error"))
			},
			expectedError: true,
			errorContains: "failed to apply header styles",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			mockOps := NewMockTableOperations(ctrl)
			tt.setupMock(mockOps)

			// Use the mock directly - no wrapper needed!
			err := tt.table.renderStyles(mockOps)

			if tt.expectedError {
				if err == nil {
					t.Errorf("Expected error but got none")
				} else if tt.errorContains != "" && !stringContains(err.Error(), tt.errorContains) {
					t.Errorf("Expected error to contain '%s', got: %s", tt.errorContains, err.Error())
				}
			} else {
				if err != nil {
					t.Errorf("Expected no error but got: %s", err.Error())
				}
			}
		})
	}
}

func TestTable_applyHeaderStyles(t *testing.T) {
	ctrl := gomock.NewController(t)
	defer ctrl.Finish()

	tests := []struct {
		name          string
		table         *Table
		setupMock     func(*MockTableOperations)
		expectedError bool
		errorContains string
	}{
		{
			name: "single_level_headers",
			table: &Table{
				Columns: Columns{
					{Name: "name", Label: "Name"},
					{Name: "age", Label: "Age"},
				},
				WriteHeader: true,
			},
			setupMock: func(mockOps *MockTableOperations) {
				headerStyle := Style{
					Bold:            true,
					BackgroundColor: "#E0E0E0",
					Alignment:       AlignmentCenterMiddle,
				}
				mockOps.EXPECT().applyStyleToRange(1, 1, 2, 1, headerStyle).Return(nil)

				// Header borders
				border := &Border{Style: BorderStyleThin}
				for col := 1; col <= 2; col++ {
					for _, side := range []string{"left", "right", "top", "bottom"} {
						mockOps.EXPECT().applyBorderToCell(col, 1, side, border).Return(nil)
					}
				}
			},
			expectedError: false,
		},
		{
			name: "multi_level_headers",
			table: &Table{
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
				WriteHeader: true,
			},
			setupMock: func(mockOps *MockTableOperations) {
				headerStyle := Style{
					Bold:            true,
					BackgroundColor: "#E0E0E0",
					Alignment:       AlignmentCenterMiddle,
				}
				mockOps.EXPECT().applyStyleToRange(1, 1, 2, 2, headerStyle).Return(nil)

				// Header borders for 2-level headers
				border := &Border{Style: BorderStyleThin}
				for row := 1; row <= 2; row++ {
					for col := 1; col <= 2; col++ {
						for _, side := range []string{"left", "right", "top", "bottom"} {
							mockOps.EXPECT().applyBorderToCell(col, row, side, border).Return(nil)
						}
					}
				}
			},
			expectedError: false,
		},
		{
			name: "header_style_error",
			table: &Table{
				Columns: Columns{
					{Name: "name", Label: "Name"},
				},
				WriteHeader: true,
			},
			setupMock: func(mockOps *MockTableOperations) {
				// First, borders are applied to all header cells
				border := &Border{Style: BorderStyleThin}
				mockOps.EXPECT().applyBorderToCell(1, 1, "left", border).Return(nil)
				mockOps.EXPECT().applyBorderToCell(1, 1, "right", border).Return(nil)
				mockOps.EXPECT().applyBorderToCell(1, 1, "top", border).Return(nil)
				mockOps.EXPECT().applyBorderToCell(1, 1, "bottom", border).Return(nil)

				// Then the style call fails
				headerStyle := Style{
					Bold:            true,
					BackgroundColor: "#E0E0E0",
					Alignment:       AlignmentCenterMiddle,
				}
				mockOps.EXPECT().applyStyleToRange(1, 1, 1, 1, headerStyle).Return(errors.New("style error"))
			},
			expectedError: true,
			errorContains: "failed to apply header cell styles",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			mockOps := NewMockTableOperations(ctrl)
			tt.setupMock(mockOps)

			err := tt.table.applyHeaderStyles(mockOps)

			if tt.expectedError {
				if err == nil {
					t.Errorf("Expected error but got none")
				} else if tt.errorContains != "" && !stringContains(err.Error(), tt.errorContains) {
					t.Errorf("Expected error to contain '%s', got: %s", tt.errorContains, err.Error())
				}
			} else {
				if err != nil {
					t.Errorf("Expected no error but got: %s", err.Error())
				}
			}
		})
	}
}

func TestTable_applyHeaderCellStyles(t *testing.T) {
	ctrl := gomock.NewController(t)
	defer ctrl.Finish()

	tests := []struct {
		name          string
		table         *Table
		setupMock     func(*MockTableOperations)
		expectedError bool
	}{
		{
			name: "success_basic_headers",
			table: &Table{
				Columns: Columns{
					{Name: "name", Label: "Name"},
					{Name: "age", Label: "Age"},
				},
			},
			setupMock: func(mockOps *MockTableOperations) {
				expectedStyle := Style{
					Bold:            true,
					BackgroundColor: "#E0E0E0",
					Alignment:       AlignmentCenterMiddle,
				}
				mockOps.EXPECT().applyStyleToRange(1, 1, 2, 1, expectedStyle).Return(nil)
			},
			expectedError: false,
		},
		{
			name: "error_applying_style",
			table: &Table{
				Columns: Columns{
					{Name: "name", Label: "Name"},
				},
			},
			setupMock: func(mockOps *MockTableOperations) {
				expectedStyle := Style{
					Bold:            true,
					BackgroundColor: "#E0E0E0",
					Alignment:       AlignmentCenterMiddle,
				}
				mockOps.EXPECT().applyStyleToRange(1, 1, 1, 1, expectedStyle).Return(errors.New("style error"))
			},
			expectedError: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			mockOps := NewMockTableOperations(ctrl)
			tt.setupMock(mockOps)

			err := tt.table.applyHeaderCellStyles(mockOps)

			if tt.expectedError && err == nil {
				t.Errorf("Expected error but got none")
			} else if !tt.expectedError && err != nil {
				t.Errorf("Expected no error but got: %s", err.Error())
			}
		})
	}
}

func TestTable_applyCellStyles(t *testing.T) {
	ctrl := gomock.NewController(t)
	defer ctrl.Finish()

	tests := []struct {
		name      string
		table     *Table
		startRow  int
		endRow    int
		setupMock func(*MockTableOperations)
	}{
		{
			name: "cell_style_priority_over_row_and_column",
			table: &Table{
				Data: DataSlice{
					{"name": "John", "age": 30},
				},
				Columns: Columns{
					{
						Name:  "name",
						Label: "Name",
						Style: &Style{Bold: true}, // Column style
					},
					{Name: "age", Label: "Age"},
				},
				RowOptionsMap: RowOptionsMap{
					0: RowOptions{
						Style: &Style{Italic: true}, // Row style
					},
				},
				CellOptionsMap: CellOptionsMap{
					1: map[int]CellOptions{
						0: {
							Style: &Style{Underline: "single"}, // Cell style - highest priority
						},
					},
				},
				WriteHeader: true,
			},
			startRow: 2,
			endRow:   2,
			setupMock: func(mockOps *MockTableOperations) {
				// Cell style should take priority
				cellStyle := Style{Underline: "single"}
				mockOps.EXPECT().applyStyleToCell(1, 2, cellStyle).Return(nil)

				// Row style should be used for second column
				rowStyle := Style{Italic: true}
				mockOps.EXPECT().applyStyleToCell(2, 2, rowStyle).Return(nil)
			},
		},
		{
			name: "row_style_priority_over_column",
			table: &Table{
				Data: DataSlice{
					{"name": "John", "age": 30},
				},
				Columns: Columns{
					{
						Name:  "name",
						Label: "Name",
						Style: &Style{Bold: true}, // Column style
					},
					{
						Name:  "age",
						Label: "Age",
						Style: &Style{FontSize: 12}, // Column style
					},
				},
				RowOptionsMap: RowOptionsMap{
					0: RowOptions{
						Style: &Style{Italic: true}, // Row style - should take priority
					},
				},
				WriteHeader: true,
			},
			startRow: 2,
			endRow:   2,
			setupMock: func(mockOps *MockTableOperations) {
				rowStyle := Style{Italic: true}
				mockOps.EXPECT().applyStyleToCell(1, 2, rowStyle).Return(nil)
				mockOps.EXPECT().applyStyleToCell(2, 2, rowStyle).Return(nil)
			},
		},
		{
			name: "column_style_only",
			table: &Table{
				Data: DataSlice{
					{"name": "John", "age": 30},
				},
				Columns: Columns{
					{
						Name:  "name",
						Label: "Name",
						Style: &Style{Bold: true},
					},
					{Name: "age", Label: "Age"}, // No style
				},
				WriteHeader: true,
			},
			startRow: 2,
			endRow:   2,
			setupMock: func(mockOps *MockTableOperations) {
				columnStyle := Style{Bold: true}
				mockOps.EXPECT().applyStyleToCell(1, 2, columnStyle).Return(nil)
				// No call expected for second column (no style)
			},
		},
		{
			name: "no_styles",
			table: &Table{
				Data: DataSlice{
					{"name": "John"},
				},
				Columns: Columns{
					{Name: "name", Label: "Name"},
				},
				WriteHeader: true,
			},
			startRow: 2,
			endRow:   2,
			setupMock: func(mockOps *MockTableOperations) {
				// No style calls expected
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			mockOps := NewMockTableOperations(ctrl)
			tt.setupMock(mockOps)

			err := tt.table.applyCellStyles(tt.startRow, tt.endRow, mockOps)
			if err != nil {
				t.Errorf("Expected no error but got: %s", err.Error())
			}
		})
	}
}

func TestTable_applyCellStyle(t *testing.T) {
	ctrl := gomock.NewController(t)
	defer ctrl.Finish()

	tests := []struct {
		name          string
		style         *Style
		colIndex      int
		rowIndex      int
		setupMock     func(*MockTableOperations)
		expectedError bool
	}{
		{
			name: "apply_valid_style",
			style: &Style{
				Bold:      true,
				TextColor: "#FF0000",
			},
			colIndex:      1,
			rowIndex:      2,
			expectedError: false,
			setupMock: func(mockOps *MockTableOperations) {
				expectedStyle := Style{
					Bold:      true,
					TextColor: "#FF0000",
				}
				mockOps.EXPECT().applyStyleToCell(1, 2, expectedStyle).Return(nil)
			},
		},
		{
			name:          "nil_style",
			style:         nil,
			colIndex:      1,
			rowIndex:      2,
			expectedError: false,
			setupMock: func(mockOps *MockTableOperations) {
				// No calls expected for nil style
			},
		},
		{
			name: "error_applying_style",
			style: &Style{
				Bold: true,
			},
			colIndex:      1,
			rowIndex:      2,
			expectedError: true,
			setupMock: func(mockOps *MockTableOperations) {
				expectedStyle := Style{Bold: true}
				mockOps.EXPECT().applyStyleToCell(1, 2, expectedStyle).Return(errors.New("apply error"))
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			mockOps := NewMockTableOperations(ctrl)
			tt.setupMock(mockOps)

			table := &Table{}
			err := table.applyCellStyle(tt.style, tt.colIndex, tt.rowIndex, mockOps)

			if tt.expectedError && err == nil {
				t.Errorf("Expected error but got none")
			} else if !tt.expectedError && err != nil {
				t.Errorf("Expected no error but got: %s", err.Error())
			}
		})
	}
}

func TestTable_applyColumnBorders(t *testing.T) {
	ctrl := gomock.NewController(t)
	defer ctrl.Finish()

	tests := []struct {
		name         string
		table        *Table
		dataStartRow int
		dataEndRow   int
		setupMock    func(*MockTableOperations)
	}{
		{
			name: "column_with_inner_borders",
			table: &Table{
				Columns: Columns{
					{
						Name:  "name",
						Label: "Name",
						Borders: &Borders{
							Left:   &Border{Style: BorderStyleThin},
							Right:  &Border{Style: BorderStyleThin},
							Top:    &Border{Style: BorderStyleThin},
							Bottom: &Border{Style: BorderStyleThin},
							Inner: &Borders{
								Left:   &Border{Style: BorderStyleThin},
								Right:  &Border{Style: BorderStyleThin},
								Top:    &Border{Style: BorderStyleThin},
								Bottom: &Border{Style: BorderStyleThin},
							},
						},
					},
				},
			},
			dataStartRow: 2,
			dataEndRow:   3,
			setupMock: func(mockOps *MockTableOperations) {
				border := &Border{Style: BorderStyleThin}
				// Inner borders applied to all cells (2 rows * 4 sides = 8 calls)
				for row := 2; row <= 3; row++ {
					for _, side := range []string{"left", "right", "top", "bottom"} {
						mockOps.EXPECT().applyBorderToCell(1, row, side, border).Return(nil)
					}
				}
			},
		},
		{
			name: "column_with_boundary_borders",
			table: &Table{
				Columns: Columns{
					{
						Name:  "name",
						Label: "Name",
						Borders: &Borders{
							Left:   &Border{Style: BorderStyleThin},
							Right:  &Border{Style: BorderStyleThin},
							Top:    &Border{Style: BorderStyleThick},
							Bottom: &Border{Style: BorderStyleThick},
						},
					},
				},
			},
			dataStartRow: 2,
			dataEndRow:   3,
			setupMock: func(mockOps *MockTableOperations) {
				leftRightBorder := &Border{Style: BorderStyleThin}
				topBottomBorder := &Border{Style: BorderStyleThick}

				// Left/right borders for all rows
				for row := 2; row <= 3; row++ {
					mockOps.EXPECT().applyBorderToCell(1, row, "left", leftRightBorder).Return(nil)
					mockOps.EXPECT().applyBorderToCell(1, row, "right", leftRightBorder).Return(nil)
				}
				// Top border only for first row
				mockOps.EXPECT().applyBorderToCell(1, 2, "top", topBottomBorder).Return(nil)
				// Bottom border only for last row
				mockOps.EXPECT().applyBorderToCell(1, 3, "bottom", topBottomBorder).Return(nil)
			},
		},
		{
			name: "column_without_borders",
			table: &Table{
				Columns: Columns{
					{Name: "name", Label: "Name"}, // No borders
				},
			},
			dataStartRow: 2,
			dataEndRow:   3,
			setupMock: func(mockOps *MockTableOperations) {
				// No border calls expected
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			mockOps := NewMockTableOperations(ctrl)
			tt.setupMock(mockOps)

			err := tt.table.applyColumnBorders(tt.dataStartRow, tt.dataEndRow, mockOps)
			if err != nil {
				t.Errorf("Expected no error but got: %s", err.Error())
			}
		})
	}
}

func TestTable_applyRowBorders(t *testing.T) {
	ctrl := gomock.NewController(t)
	defer ctrl.Finish()

	tests := []struct {
		name         string
		table        *Table
		dataRowIndex int
		actualRowNum int
		totalColumns int
		setupMock    func(*MockTableOperations)
	}{
		{
			name: "row_with_inner_borders",
			table: &Table{
				RowOptionsMap: RowOptionsMap{
					0: RowOptions{
						Border: &Borders{
							Left:   &Border{Style: BorderStyleThin},
							Right:  &Border{Style: BorderStyleThin},
							Top:    &Border{Style: BorderStyleThin},
							Bottom: &Border{Style: BorderStyleThin},
							Inner: &Borders{
								Left:   &Border{Style: BorderStyleThin},
								Right:  &Border{Style: BorderStyleThin},
								Top:    &Border{Style: BorderStyleThin},
								Bottom: &Border{Style: BorderStyleThin},
							},
						},
					},
				},
			},
			dataRowIndex: 0,
			actualRowNum: 2,
			totalColumns: 2,
			setupMock: func(mockOps *MockTableOperations) {
				border := &Border{Style: BorderStyleThin}
				// Inner borders applied to all cells in row (2 columns * 4 sides = 8 calls)
				for col := 1; col <= 2; col++ {
					for _, side := range []string{"left", "right", "top", "bottom"} {
						mockOps.EXPECT().applyBorderToCell(col, 2, side, border).Return(nil)
					}
				}
			},
		},
		{
			name: "row_with_boundary_borders",
			table: &Table{
				RowOptionsMap: RowOptionsMap{
					0: RowOptions{
						Border: &Borders{
							Left:   &Border{Style: BorderStyleThin},
							Right:  &Border{Style: BorderStyleThin},
							Top:    &Border{Style: BorderStyleThick},
							Bottom: &Border{Style: BorderStyleThick},
						},
					},
				},
			},
			dataRowIndex: 0,
			actualRowNum: 2,
			totalColumns: 2,
			setupMock: func(mockOps *MockTableOperations) {
				leftRightBorder := &Border{Style: BorderStyleThin}
				topBottomBorder := &Border{Style: BorderStyleThick}

				// Top/bottom borders for all columns
				for col := 1; col <= 2; col++ {
					mockOps.EXPECT().applyBorderToCell(col, 2, "top", topBottomBorder).Return(nil)
					mockOps.EXPECT().applyBorderToCell(col, 2, "bottom", topBottomBorder).Return(nil)
				}
				// Left border only for first column
				mockOps.EXPECT().applyBorderToCell(1, 2, "left", leftRightBorder).Return(nil)
				// Right border only for last column
				mockOps.EXPECT().applyBorderToCell(2, 2, "right", leftRightBorder).Return(nil)
			},
		},
		{
			name: "row_without_borders",
			table: &Table{
				RowOptionsMap: RowOptionsMap{},
			},
			dataRowIndex: 0,
			actualRowNum: 2,
			totalColumns: 2,
			setupMock: func(mockOps *MockTableOperations) {
				// No border calls expected
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			mockOps := NewMockTableOperations(ctrl)
			tt.setupMock(mockOps)

			err := tt.table.applyRowBorders(tt.dataRowIndex, tt.actualRowNum, tt.totalColumns, mockOps)
			if err != nil {
				t.Errorf("Expected no error but got: %s", err.Error())
			}
		})
	}
}

func TestTable_applyCellSpecificBorders(t *testing.T) {
	ctrl := gomock.NewController(t)
	defer ctrl.Finish()

	tests := []struct {
		name         string
		table        *Table
		dataStartRow int
		setupMock    func(*MockTableOperations)
	}{
		{
			name: "apply_cell_specific_borders",
			table: &Table{
				CellOptionsMap: CellOptionsMap{
					1: map[int]CellOptions{
						0: {
							Border: &Borders{
								Left:   &Border{Style: BorderStyleThin},
								Right:  &Border{Style: BorderStyleThick},
								Top:    &Border{Style: BorderStyleDotted},
								Bottom: &Border{Style: BorderStyleDashed},
							},
						},
					},
					2: map[int]CellOptions{
						1: {
							Border: &Borders{
								Left: &Border{Style: BorderStyleThin},
							},
						},
					},
				},
			},
			dataStartRow: 2,
			setupMock: func(mockOps *MockTableOperations) {
				// First cell borders (4 borders)
				mockOps.EXPECT().applyBorderToCell(1, 2, "left", &Border{Style: BorderStyleThin}).Return(nil)
				mockOps.EXPECT().applyBorderToCell(1, 2, "right", &Border{Style: BorderStyleThick}).Return(nil)
				mockOps.EXPECT().applyBorderToCell(1, 2, "top", &Border{Style: BorderStyleDotted}).Return(nil)
				mockOps.EXPECT().applyBorderToCell(1, 2, "bottom", &Border{Style: BorderStyleDashed}).Return(nil)

				// Second cell borders (1 border)
				mockOps.EXPECT().applyBorderToCell(2, 3, "left", &Border{Style: BorderStyleThin}).Return(nil)
			},
		},
		{
			name: "no_cell_borders",
			table: &Table{
				CellOptionsMap: CellOptionsMap{},
			},
			dataStartRow: 2,
			setupMock: func(mockOps *MockTableOperations) {
				// No border calls expected
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			mockOps := NewMockTableOperations(ctrl)
			tt.setupMock(mockOps)

			err := tt.table.applyCellSpecificBorders(tt.dataStartRow, mockOps)
			if err != nil {
				t.Errorf("Expected no error but got: %s", err.Error())
			}
		})
	}
}

func TestTable_applyBordersToCell(t *testing.T) {
	ctrl := gomock.NewController(t)
	defer ctrl.Finish()

	tests := []struct {
		name          string
		col           int
		row           int
		borders       *Borders
		setupMock     func(*MockTableOperations)
		expectedError bool
		errorContains string
	}{
		{
			name: "apply_all_borders",
			col:  1,
			row:  2,
			borders: &Borders{
				Left:   &Border{Style: BorderStyleThin},
				Right:  &Border{Style: BorderStyleThick},
				Top:    &Border{Style: BorderStyleDotted},
				Bottom: &Border{Style: BorderStyleDashed},
			},
			setupMock: func(mockOps *MockTableOperations) {
				mockOps.EXPECT().applyBorderToCell(1, 2, "left", &Border{Style: BorderStyleThin}).Return(nil)
				mockOps.EXPECT().applyBorderToCell(1, 2, "right", &Border{Style: BorderStyleThick}).Return(nil)
				mockOps.EXPECT().applyBorderToCell(1, 2, "top", &Border{Style: BorderStyleDotted}).Return(nil)
				mockOps.EXPECT().applyBorderToCell(1, 2, "bottom", &Border{Style: BorderStyleDashed}).Return(nil)
			},
			expectedError: false,
		},
		{
			name: "apply_partial_borders",
			col:  1,
			row:  2,
			borders: &Borders{
				Left:  &Border{Style: BorderStyleThin},
				Right: &Border{Style: BorderStyleThick},
				// Top and Bottom are nil
			},
			setupMock: func(mockOps *MockTableOperations) {
				mockOps.EXPECT().applyBorderToCell(1, 2, "left", &Border{Style: BorderStyleThin}).Return(nil)
				mockOps.EXPECT().applyBorderToCell(1, 2, "right", &Border{Style: BorderStyleThick}).Return(nil)
			},
			expectedError: false,
		},
		{
			name: "error_applying_left_border",
			col:  1,
			row:  2,
			borders: &Borders{
				Left: &Border{Style: BorderStyleThin},
			},
			setupMock: func(mockOps *MockTableOperations) {
				mockOps.EXPECT().applyBorderToCell(1, 2, "left", &Border{Style: BorderStyleThin}).Return(errors.New("left border error"))
			},
			expectedError: true,
			errorContains: "failed to apply left border",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			mockOps := NewMockTableOperations(ctrl)
			tt.setupMock(mockOps)

			table := &Table{}
			err := table.applyBordersToCell(tt.col, tt.row, tt.borders, mockOps)

			if tt.expectedError {
				if err == nil {
					t.Errorf("Expected error but got none")
				} else if tt.errorContains != "" && !stringContains(err.Error(), tt.errorContains) {
					t.Errorf("Expected error to contain '%s', got: %s", tt.errorContains, err.Error())
				}
			} else {
				if err != nil {
					t.Errorf("Expected no error but got: %s", err.Error())
				}
			}
		})
	}
}

// Helper function to check if a string contains a substring
func stringContains(s, substr string) bool {
	if len(substr) == 0 {
		return true
	}
	if len(s) < len(substr) {
		return false
	}
	for i := 0; i <= len(s)-len(substr); i++ {
		if s[i:i+len(substr)] == substr {
			return true
		}
	}
	return false
}
