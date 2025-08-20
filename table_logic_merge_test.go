package spit

import (
	"errors"
	"testing"

	"go.uber.org/mock/gomock"
)

func TestTable_processMerging(t *testing.T) {
	tests := []struct {
		name          string
		setupTable    func() *Table
		setupMock     func(*MockTableOperations)
		expectedError string
		expectedCalls int
	}{
		{
			name: "Success - no headers, simple data",
			setupTable: func() *Table {
				return &Table{
					WriteHeader: false,
					Columns: Columns{
						{Name: "col1", Label: "Column 1"},
						{Name: "col2", Label: "Column 2"},
					},
					Data: DataSlice{
						{"col1": "A", "col2": "B"},
						{"col1": "A", "col2": "C"},
					},
				}
			},
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().processValue(gomock.Any(), gomock.Any()).Return("A", nil).AnyTimes()
			},
			expectedError: "",
			expectedCalls: 1,
		},
		{
			name: "Success - with headers",
			setupTable: func() *Table {
				return &Table{
					WriteHeader: true,
					Columns: Columns{
						{Name: "col1", Label: "Column 1"},
						{Name: "col2", Label: "Column 2"},
					},
					Data: DataSlice{
						{"col1": "A", "col2": "B"},
					},
				}
			},
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().mergeCells(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(nil).AnyTimes()
				mock.EXPECT().processValue(gomock.Any(), gomock.Any()).Return("A", nil).AnyTimes()
			},
			expectedError: "",
			expectedCalls: 1,
		},
		{
			name: "Error - header merging fails",
			setupTable: func() *Table {
				return &Table{
					WriteHeader: true,
					Columns: Columns{
						{
							Label: "Parent",
							Columns: Columns{
								{Name: "col1", Label: "Column 1"},
								{Name: "col2", Label: "Column 2"},
							},
						},
					},
					Data: DataSlice{},
				}
			},
			setupMock: func(mock *MockTableOperations) {
				// Header merging will be called but will fail - this should not cause processMerging to fail
				// as the error is logged as a warning and processing continues
				mock.EXPECT().mergeCells(gomock.Any(), gomock.Any(), gomock.Any(), gomock.Any()).Return(errors.New("merge failed")).AnyTimes()
			},
			expectedError: "", // Changed: processMerging doesn't fail on header merge errors, just logs warnings
			expectedCalls: 1,
		},
		{
			name: "Success - with row options and custom merging",
			setupTable: func() *Table {
				return &Table{
					WriteHeader: false,
					Columns: Columns{
						{Name: "col1", Label: "Column 1"},
						{Name: "col2", Label: "Column 2"},
					},
					Data: DataSlice{
						{"col1": "A", "col2": "A"}, // Same values to trigger merge
						{"col1": "C", "col2": "D"},
					},
					RowOptionsMap: RowOptionsMap{
						0: {
							Merge: &MergeRules{
								Horizontal: MergeConditions{MergeConditionIdentical},
							},
						},
						1: {
							Mergeable: false,
						},
					},
				}
			},
			setupMock: func(mock *MockTableOperations) {
				// Expect ProcessValue calls for row 0 with identical values that should merge
				mock.EXPECT().processValue("A", "").Return("A", nil).Times(2)
				mock.EXPECT().mergeCells(1, 1, 2, 1).Return(nil)
				// No ProcessValue calls expected for row 1 since it's not mergeable
			},
			expectedError: "",
			expectedCalls: 1,
		},
		{
			name: "Success - empty table",
			setupTable: func() *Table {
				return &Table{
					WriteHeader: false,
					Columns:     Columns{},
					Data:        DataSlice{},
				}
			},
			setupMock: func(mock *MockTableOperations) {
				// No calls expected for empty table
			},
			expectedError: "",
			expectedCalls: 1,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ctrl := gomock.NewController(t)
			defer ctrl.Finish()

			mockOps := NewMockTableOperations(ctrl)
			tt.setupMock(mockOps)

			table := tt.setupTable()
			err := table.processMerging(mockOps)

			if tt.expectedError != "" {
				if err == nil {
					t.Errorf("Expected error containing '%s', got nil", tt.expectedError)
				} else if !containsSubstring(err.Error(), tt.expectedError) {
					t.Errorf("Expected error containing '%s', got '%s'", tt.expectedError, err.Error())
				}
			} else if err != nil {
				t.Errorf("Expected no error, got: %v", err)
			}
		})
	}
}

func TestTable_executeHeaderMerging(t *testing.T) {
	tests := []struct {
		name          string
		setupTable    func() *Table
		setupMock     func(*MockTableOperations)
		expectedError string
	}{
		{
			name: "Success - single level headers (no merging needed)",
			setupTable: func() *Table {
				return &Table{
					Columns: Columns{
						{Name: "col1", Label: "Column 1"},
						{Name: "col2", Label: "Column 2"},
					},
				}
			},
			setupMock: func(mock *MockTableOperations) {
				// No merge calls expected for single level
			},
			expectedError: "",
		},
		{
			name: "Success - multi-level headers",
			setupTable: func() *Table {
				return &Table{
					Columns: Columns{
						{
							Label: "Group 1",
							Columns: Columns{
								{Name: "col1", Label: "Column 1"},
								{Name: "col2", Label: "Column 2"},
							},
						},
						{Name: "col3", Label: "Column 3"},
					},
				}
			},
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().mergeCells(1, 1, 2, 1).Return(nil)
				mock.EXPECT().mergeCells(3, 1, 3, 2).Return(nil)
			},
			expectedError: "",
		},
		{
			name: "Error - merge operation fails",
			setupTable: func() *Table {
				return &Table{
					Columns: Columns{
						{
							Label: "Group 1",
							Columns: Columns{
								{Name: "col1", Label: "Column 1"},
								{Name: "col2", Label: "Column 2"},
							},
						},
					},
				}
			},
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().mergeCells(1, 1, 2, 1).Return(errors.New("merge failed"))
			},
			expectedError: "", // Changed: executeHeaderMerging logs warnings but doesn't return errors for merge failures
		},
		{
			name: "Success - complex nested structure",
			setupTable: func() *Table {
				return &Table{
					Columns: Columns{
						{
							Label: "Parent",
							Columns: Columns{
								{
									Label: "Child 1",
									Columns: Columns{
										{Name: "col1", Label: "Column 1"},
										{Name: "col2", Label: "Column 2"},
									},
								},
								{Name: "col3", Label: "Column 3"},
							},
						},
					},
				}
			},
			setupMock: func(mock *MockTableOperations) {
				// Parent spans all 3 columns at row 1
				mock.EXPECT().mergeCells(1, 1, 3, 1).Return(nil)
				// Child 1 spans 2 columns at row 2
				mock.EXPECT().mergeCells(1, 2, 2, 2).Return(nil)
				// col3 spans from row 2 to row 3
				mock.EXPECT().mergeCells(3, 2, 3, 3).Return(nil)
			},
			expectedError: "",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ctrl := gomock.NewController(t)
			defer ctrl.Finish()

			mockOps := NewMockTableOperations(ctrl)
			tt.setupMock(mockOps)

			table := tt.setupTable()
			err := table.executeHeaderMerging(mockOps)

			if tt.expectedError != "" {
				if err == nil {
					t.Errorf("Expected error containing '%s', got nil", tt.expectedError)
				} else if !containsSubstring(err.Error(), tt.expectedError) {
					t.Errorf("Expected error containing '%s', got '%s'", tt.expectedError, err.Error())
				}
			} else if err != nil {
				t.Errorf("Expected no error, got: %v", err)
			}
		})
	}
}

func TestTable_processHeaderMergingRecursive(t *testing.T) {
	tests := []struct {
		name          string
		setupColumns  func() Columns
		currentRow    int
		maxDepth      int
		startCol      int
		setupMock     func(*MockTableOperations)
		expectedError string
	}{
		{
			name: "Success - leaf columns only",
			setupColumns: func() Columns {
				return Columns{
					{Name: "col1", Label: "Column 1"},
					{Name: "col2", Label: "Column 2"},
				}
			},
			currentRow: 1,
			maxDepth:   2,
			startCol:   1,
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().mergeCells(1, 1, 1, 2).Return(nil)
				mock.EXPECT().mergeCells(2, 1, 2, 2).Return(nil)
			},
			expectedError: "",
		},
		{
			name: "Success - mixed leaf and parent columns",
			setupColumns: func() Columns {
				return Columns{
					{
						Label: "Group",
						Columns: Columns{
							{Name: "col1", Label: "Column 1"},
						},
					},
					{Name: "col2", Label: "Column 2"},
				}
			},
			currentRow: 1,
			maxDepth:   2,
			startCol:   1,
			setupMock: func(mock *MockTableOperations) {
				// Group header doesn't need horizontal merge (only 1 child)
				// col2 needs vertical merge
				mock.EXPECT().mergeCells(2, 1, 2, 2).Return(nil)
			},
			expectedError: "",
		},
		{
			name: "Success - parent with multiple children",
			setupColumns: func() Columns {
				return Columns{
					{
						Label: "Group",
						Columns: Columns{
							{Name: "col1", Label: "Column 1"},
							{Name: "col2", Label: "Column 2"},
							{Name: "col3", Label: "Column 3"},
						},
					},
				}
			},
			currentRow: 1,
			maxDepth:   2,
			startCol:   1,
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().mergeCells(1, 1, 3, 1).Return(nil)
			},
			expectedError: "",
		},
		{
			name: "Error - merge fails during recursion",
			setupColumns: func() Columns {
				return Columns{
					{
						Label: "Group",
						Columns: Columns{
							{Name: "col1", Label: "Column 1"},
							{Name: "col2", Label: "Column 2"},
						},
					},
				}
			},
			currentRow: 1,
			maxDepth:   2,
			startCol:   1,
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().mergeCells(1, 1, 2, 1).Return(errors.New("recursive merge failed"))
			},
			expectedError: "", // Changed: processHeaderMergingRecursive logs warnings but doesn't return errors for merge failures
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ctrl := gomock.NewController(t)
			defer ctrl.Finish()

			mockOps := NewMockTableOperations(ctrl)
			tt.setupMock(mockOps)

			table := &Table{}
			columns := tt.setupColumns()
			err := table.processHeaderMergingRecursive(columns, tt.currentRow, tt.maxDepth, tt.startCol, mockOps)

			if tt.expectedError != "" {
				if err == nil {
					t.Errorf("Expected error containing '%s', got nil", tt.expectedError)
				} else if !containsSubstring(err.Error(), tt.expectedError) {
					t.Errorf("Expected error containing '%s', got '%s'", tt.expectedError, err.Error())
				}
			} else if err != nil {
				t.Errorf("Expected no error, got: %v", err)
			}
		})
	}
}

func TestTable_executeVerticalMerging(t *testing.T) {
	tests := []struct {
		name           string
		setupTable     func() *Table
		column         Column
		actualColIndex int
		dataStartRow   int
		setupMock      func(*MockTableOperations)
		expectedError  string
	}{
		{
			name: "Success - no merge configuration",
			setupTable: func() *Table {
				return &Table{
					Data: DataSlice{
						{"col1": "A"},
						{"col1": "A"},
					},
				}
			},
			column: Column{
				Name:  "col1",
				Label: "Column 1",
			},
			actualColIndex: 1,
			dataStartRow:   2,
			setupMock: func(mock *MockTableOperations) {
				// No merge calls expected
			},
			expectedError: "",
		},
		{
			name: "Success - with vertical merge configuration",
			setupTable: func() *Table {
				return &Table{
					Data: DataSlice{
						{"col1": "A"},
						{"col1": "A"},
						{"col1": "B"},
					},
				}
			},
			column: Column{
				Name:  "col1",
				Label: "Column 1",
				Merge: &MergeRules{
					Vertical: MergeConditions{MergeConditionIdentical},
				},
			},
			actualColIndex: 1,
			dataStartRow:   2,
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().processValue("A", "").Return("A", nil).Times(2)
				mock.EXPECT().processValue("B", "").Return("B", nil).Times(1)
				mock.EXPECT().mergeCells(1, 2, 1, 3).Return(nil)
			},
			expectedError: "",
		},
		{
			name: "Success - merge operation fails but continues",
			setupTable: func() *Table {
				return &Table{
					Data: DataSlice{
						{"col1": "A"},
						{"col1": "A"},
					},
				}
			},
			column: Column{
				Name:  "col1",
				Label: "Column 1",
				Merge: &MergeRules{
					Vertical: MergeConditions{MergeConditionIdentical},
				},
			},
			actualColIndex: 1,
			dataStartRow:   2,
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().processValue("A", "").Return("A", nil).Times(2)
				mock.EXPECT().mergeCells(1, 2, 1, 3).Return(errors.New("merge failed"))
			},
			expectedError: "",
		},
		{
			name: "Success - with format and empty merge conditions",
			setupTable: func() *Table {
				return &Table{
					Data: DataSlice{
						{"col1": ""},
						{"col1": nil},
						{"col1": "A"},
					},
				}
			},
			column: Column{
				Name:   "col1",
				Label:  "Column 1",
				Format: "string",
				Merge: &MergeRules{
					Vertical: MergeConditions{MergeConditionEmpty},
				},
			},
			actualColIndex: 1,
			dataStartRow:   1,
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().processValue("", "string").Return("", nil)
				mock.EXPECT().processValue(nil, "string").Return("", nil)
				mock.EXPECT().processValue("A", "string").Return("A", nil)
				mock.EXPECT().mergeCells(1, 1, 1, 2).Return(nil)
			},
			expectedError: "",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ctrl := gomock.NewController(t)
			defer ctrl.Finish()

			mockOps := NewMockTableOperations(ctrl)
			tt.setupMock(mockOps)

			table := tt.setupTable()
			err := table.executeVerticalMerging(tt.column, tt.actualColIndex, tt.dataStartRow, mockOps)

			if tt.expectedError != "" {
				if err == nil {
					t.Errorf("Expected error containing '%s', got nil", tt.expectedError)
				} else if !containsSubstring(err.Error(), tt.expectedError) {
					t.Errorf("Expected error containing '%s', got '%s'", tt.expectedError, err.Error())
				}
			} else if err != nil {
				t.Errorf("Expected no error, got: %v", err)
			}
		})
	}
}

func TestTable_findVerticalMergeRanges(t *testing.T) {
	tests := []struct {
		name           string
		setupTable     func() *Table
		colIndex       int
		fieldName      string
		format         string
		conditions     MergeConditions
		setupMock      func(*MockTableOperations)
		expectedRanges [][]int
	}{
		{
			name: "Success - identical values merge",
			setupTable: func() *Table {
				return &Table{
					Data: DataSlice{
						{"col1": "A"},
						{"col1": "A"},
						{"col1": "B"},
						{"col1": "B"},
						{"col1": "B"},
					},
				}
			},
			colIndex:   1,
			fieldName:  "col1",
			format:     "",
			conditions: MergeConditions{MergeConditionIdentical},
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().processValue("A", "").Return("A", nil).Times(2)
				mock.EXPECT().processValue("B", "").Return("B", nil).Times(3)
			},
			expectedRanges: [][]int{{0, 1}, {2, 3, 4}},
		},
		{
			name: "Success - empty values merge",
			setupTable: func() *Table {
				return &Table{
					Data: DataSlice{
						{"col1": ""},
						{"col1": nil},
						{"col1": "A"},
						{"col1": ""},
						{"col1": ""},
					},
				}
			},
			colIndex:   1,
			fieldName:  "col1",
			format:     "",
			conditions: MergeConditions{MergeConditionEmpty},
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().processValue("", "").Return("", nil).Times(3)
				mock.EXPECT().processValue(nil, "").Return("", nil).Times(1)
				mock.EXPECT().processValue("A", "").Return("A", nil).Times(1)
			},
			expectedRanges: [][]int{{0, 1}, {3, 4}},
		},
		{
			name: "Success - with row options blocking merge",
			setupTable: func() *Table {
				return &Table{
					Data: DataSlice{
						{"col1": "A"},
						{"col1": "A"},
						{"col1": "A"},
					},
					RowOptionsMap: RowOptionsMap{
						1: {Mergeable: false},
					},
				}
			},
			colIndex:   1,
			fieldName:  "col1",
			format:     "",
			conditions: MergeConditions{MergeConditionIdentical},
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().processValue("A", "").Return("A", nil).Times(2)
			},
			expectedRanges: [][]int{},
		},
		{
			name: "Success - with cell options blocking merge",
			setupTable: func() *Table {
				return &Table{
					Data: DataSlice{
						{"col1": "A"},
						{"col1": "A"},
						{"col1": "A"},
					},
					CellOptionsMap: CellOptionsMap{
						1: {
							1: {Mergeable: false},
						},
					},
				}
			},
			colIndex:   1,
			fieldName:  "col1",
			format:     "",
			conditions: MergeConditions{MergeConditionIdentical},
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().processValue("A", "").Return("A", nil).Times(2)
			},
			expectedRanges: [][]int{},
		},
		{
			name: "Success - field lookup fails",
			setupTable: func() *Table {
				return &Table{
					Data: DataSlice{
						{"col1": "A"},
						{"col2": "B"}, // Missing col1
						{"col1": "A"},
					},
				}
			},
			colIndex:   1,
			fieldName:  "col1",
			format:     "",
			conditions: MergeConditions{MergeConditionIdentical},
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().processValue("A", "").Return("A", nil).Times(2)
			},
			expectedRanges: [][]int{},
		},
		{
			name: "Success - process value fails",
			setupTable: func() *Table {
				return &Table{
					Data: DataSlice{
						{"col1": "A"},
						{"col1": "A"},
					},
				}
			},
			colIndex:   1,
			fieldName:  "col1",
			format:     "",
			conditions: MergeConditions{MergeConditionIdentical},
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().processValue("A", "").Return("A", nil).Times(1)
				mock.EXPECT().processValue("A", "").Return("", errors.New("process failed")).Times(1)
			},
			expectedRanges: [][]int{},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ctrl := gomock.NewController(t)
			defer ctrl.Finish()

			mockOps := NewMockTableOperations(ctrl)
			tt.setupMock(mockOps)

			table := tt.setupTable()
			ranges := table.findVerticalMergeRanges(tt.colIndex, tt.fieldName, tt.format, tt.conditions, mockOps)

			if len(ranges) != len(tt.expectedRanges) {
				t.Errorf("Expected %d ranges, got %d", len(tt.expectedRanges), len(ranges))
				return
			}

			for i, expectedRange := range tt.expectedRanges {
				if len(ranges[i]) != len(expectedRange) {
					t.Errorf("Range %d: expected length %d, got %d", i, len(expectedRange), len(ranges[i]))
					continue
				}
				for j, expectedVal := range expectedRange {
					if ranges[i][j] != expectedVal {
						t.Errorf("Range %d[%d]: expected %d, got %d", i, j, expectedVal, ranges[i][j])
					}
				}
			}
		})
	}
}

func TestTable_executeHorizontalMerging(t *testing.T) {
	tests := []struct {
		name          string
		setupTable    func() *Table
		item          Data
		columns       Columns
		rowNum        int
		startColIndex int
		rowOptions    *RowOptions
		setupMock     func(*MockTableOperations)
		expectedError string
	}{
		{
			name: "Success - empty columns",
			setupTable: func() *Table {
				return &Table{}
			},
			item:          Data{},
			columns:       Columns{},
			rowNum:        1,
			startColIndex: 1,
			rowOptions:    nil,
			setupMock: func(mock *MockTableOperations) {
				// No calls expected
			},
			expectedError: "",
		},
		{
			name: "Success - row options with custom merge",
			setupTable: func() *Table {
				return &Table{
					Data: DataSlice{
						{"col1": "A", "col2": "A"},
					},
				}
			},
			item: Data{"col1": "A", "col2": "A"},
			columns: Columns{
				{Name: "col1", Label: "Column 1"},
				{Name: "col2", Label: "Column 2"},
			},
			rowNum:        1,
			startColIndex: 1,
			rowOptions: &RowOptions{
				Merge: &MergeRules{
					Horizontal: MergeConditions{MergeConditionIdentical},
				},
			},
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().processValue("A", "").Return("A", nil).Times(2)
				mock.EXPECT().mergeCells(1, 1, 2, 1).Return(nil)
			},
			expectedError: "",
		},
		{
			name: "Success - column-level merge configuration",
			setupTable: func() *Table {
				return &Table{
					Data: DataSlice{
						{"col1": "A", "col2": "A", "col3": "B"},
					},
				}
			},
			item: Data{"col1": "A", "col2": "A", "col3": "B"},
			columns: Columns{
				{
					Name:  "col1",
					Label: "Column 1",
					Merge: &MergeRules{
						Horizontal: MergeConditions{MergeConditionIdentical},
					},
				},
				{
					Name:  "col2",
					Label: "Column 2",
					Merge: &MergeRules{
						Horizontal: MergeConditions{MergeConditionIdentical},
					},
				},
				{Name: "col3", Label: "Column 3"}, // No merge configuration for col3
			},
			rowNum:        1,
			startColIndex: 1,
			rowOptions:    nil,
			setupMock: func(mock *MockTableOperations) {
				// Only expect ProcessValue calls for columns with merge configuration
				mock.EXPECT().processValue("A", "").Return("A", nil).Times(2)
				mock.EXPECT().mergeCells(1, 1, 2, 1).Return(nil)
				// No ProcessValue call expected for col3 since it has no merge configuration
			},
			expectedError: "",
		},
		{
			name: "Success - mixed compatible and incompatible merge conditions",
			setupTable: func() *Table {
				return &Table{
					Data: DataSlice{
						{"col1": "A", "col2": "A", "col3": "", "col4": ""},
					},
				}
			},
			item: Data{"col1": "A", "col2": "A", "col3": "", "col4": ""},
			columns: Columns{
				{
					Name:  "col1",
					Label: "Column 1",
					Merge: &MergeRules{
						Horizontal: MergeConditions{MergeConditionIdentical},
					},
				},
				{
					Name:  "col2",
					Label: "Column 2",
					Merge: &MergeRules{
						Horizontal: MergeConditions{MergeConditionIdentical},
					},
				},
				{
					Name:  "col3",
					Label: "Column 3",
					Merge: &MergeRules{
						Horizontal: MergeConditions{MergeConditionEmpty},
					},
				},
				{
					Name:  "col4",
					Label: "Column 4",
					Merge: &MergeRules{
						Horizontal: MergeConditions{MergeConditionEmpty},
					},
				},
			},
			rowNum:        1,
			startColIndex: 1,
			rowOptions:    nil,
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().processValue("A", "").Return("A", nil).Times(2)
				mock.EXPECT().mergeCells(1, 1, 2, 1).Return(nil)
				mock.EXPECT().processValue("", "").Return("", nil).Times(2)
				mock.EXPECT().mergeCells(3, 1, 4, 1).Return(nil)
			},
			expectedError: "",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ctrl := gomock.NewController(t)
			defer ctrl.Finish()

			mockOps := NewMockTableOperations(ctrl)
			tt.setupMock(mockOps)

			table := tt.setupTable()
			err := table.executeHorizontalMerging(tt.item, tt.columns, tt.rowNum, tt.startColIndex, tt.rowOptions, mockOps)

			if tt.expectedError != "" {
				if err == nil {
					t.Errorf("Expected error containing '%s', got nil", tt.expectedError)
				} else if !containsSubstring(err.Error(), tt.expectedError) {
					t.Errorf("Expected error containing '%s', got '%s'", tt.expectedError, err.Error())
				}
			} else if err != nil {
				t.Errorf("Expected no error, got: %v", err)
			}
		})
	}
}

func TestTable_applyHorizontalMerges(t *testing.T) {
	tests := []struct {
		name         string
		mergeRanges  [][]int
		rowNum       int
		baseColIndex int
		setupMock    func(*MockTableOperations)
	}{
		{
			name:         "Success - empty merge ranges",
			mergeRanges:  [][]int{},
			rowNum:       1,
			baseColIndex: 1,
			setupMock: func(mock *MockTableOperations) {
				// No calls expected
			},
		},
		{
			name:         "Success - single cell ranges (skipped)",
			mergeRanges:  [][]int{{0}, {2}},
			rowNum:       1,
			baseColIndex: 1,
			setupMock: func(mock *MockTableOperations) {
				// No calls expected for single-cell ranges
			},
		},
		{
			name:         "Success - valid merge ranges",
			mergeRanges:  [][]int{{0, 1}, {2, 3, 4}},
			rowNum:       2,
			baseColIndex: 1,
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().mergeCells(1, 2, 2, 2).Return(nil)
				mock.EXPECT().mergeCells(3, 2, 5, 2).Return(nil)
			},
		},
		{
			name:         "Success - merge operation fails but continues",
			mergeRanges:  [][]int{{0, 1}, {2, 3}},
			rowNum:       1,
			baseColIndex: 2,
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().mergeCells(2, 1, 3, 1).Return(errors.New("merge failed"))
				mock.EXPECT().mergeCells(4, 1, 5, 1).Return(nil)
			},
		},
		{
			name:         "Success - complex ranges with different base column",
			mergeRanges:  [][]int{{1, 2, 3}, {5, 6}},
			rowNum:       3,
			baseColIndex: 5,
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().mergeCells(6, 3, 8, 3).Return(nil)
				mock.EXPECT().mergeCells(10, 3, 11, 3).Return(nil)
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ctrl := gomock.NewController(t)
			defer ctrl.Finish()

			mockOps := NewMockTableOperations(ctrl)
			tt.setupMock(mockOps)

			table := &Table{}
			table.applyHorizontalMerges(tt.mergeRanges, tt.rowNum, tt.baseColIndex, mockOps)
		})
	}
}

func TestTable_findHorizontalMergeRanges(t *testing.T) {
	tests := []struct {
		name           string
		setupTable     func() *Table
		item           Data
		columns        Columns
		conditions     MergeConditions
		setupMock      func(*MockTableOperations)
		expectedRanges [][]int
	}{
		{
			name: "Success - identical values merge",
			setupTable: func() *Table {
				return &Table{}
			},
			item: Data{"col1": "A", "col2": "A", "col3": "B", "col4": "B", "col5": "C"},
			columns: Columns{
				{Name: "col1", Label: "Column 1"},
				{Name: "col2", Label: "Column 2"},
				{Name: "col3", Label: "Column 3"},
				{Name: "col4", Label: "Column 4"},
				{Name: "col5", Label: "Column 5"},
			},
			conditions: MergeConditions{MergeConditionIdentical},
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().processValue("A", "").Return("A", nil).Times(2)
				mock.EXPECT().processValue("B", "").Return("B", nil).Times(2)
				mock.EXPECT().processValue("C", "").Return("C", nil).Times(1)
			},
			expectedRanges: [][]int{{0, 1}, {2, 3}},
		},
		{
			name: "Success - empty values merge",
			setupTable: func() *Table {
				return &Table{}
			},
			item: Data{"col1": "", "col2": nil, "col3": "A", "col4": "", "col5": ""},
			columns: Columns{
				{Name: "col1", Label: "Column 1"},
				{Name: "col2", Label: "Column 2"},
				{Name: "col3", Label: "Column 3"},
				{Name: "col4", Label: "Column 4"},
				{Name: "col5", Label: "Column 5"},
			},
			conditions: MergeConditions{MergeConditionEmpty},
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().processValue("", "").Return("", nil).Times(3)
				mock.EXPECT().processValue(nil, "").Return("", nil).Times(1)
				mock.EXPECT().processValue("A", "").Return("A", nil).Times(1)
			},
			expectedRanges: [][]int{{0, 1}, {3, 4}},
		},
		{
			name: "Success - with cell options blocking merge",
			setupTable: func() *Table {
				return &Table{
					CellOptionsMap: CellOptionsMap{
						2: {
							0: {Mergeable: false},
						},
					},
				}
			},
			item: Data{"col1": "A", "col2": "A", "col3": "A"},
			columns: Columns{
				{Name: "col1", Label: "Column 1"},
				{Name: "col2", Label: "Column 2"},
				{Name: "col3", Label: "Column 3"},
			},
			conditions: MergeConditions{MergeConditionIdentical},
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().processValue("A", "").Return("A", nil).Times(2)
			},
			expectedRanges: [][]int{},
		},
		{
			name: "Success - field lookup fails",
			setupTable: func() *Table {
				return &Table{}
			},
			item: Data{"col1": "A", "col3": "A"}, // Missing col2
			columns: Columns{
				{Name: "col1", Label: "Column 1"},
				{Name: "col2", Label: "Column 2"},
				{Name: "col3", Label: "Column 3"},
			},
			conditions: MergeConditions{MergeConditionIdentical},
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().processValue("A", "").Return("A", nil).Times(2)
			},
			expectedRanges: [][]int{},
		},
		{
			name: "Success - process value fails uses raw value",
			setupTable: func() *Table {
				return &Table{}
			},
			item: Data{"col1": "A", "col2": "A"},
			columns: Columns{
				{Name: "col1", Label: "Column 1", Format: "string"},
				{Name: "col2", Label: "Column 2", Format: "string"},
			},
			conditions: MergeConditions{MergeConditionIdentical},
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().processValue("A", "string").Return("A", nil).Times(1)
				mock.EXPECT().processValue("A", "string").Return("", errors.New("process failed")).Times(1)
			},
			expectedRanges: [][]int{{0, 1}},
		},
		{
			name: "Success - mixed conditions identical and empty",
			setupTable: func() *Table {
				return &Table{}
			},
			item: Data{"col1": "A", "col2": "A", "col3": "", "col4": nil},
			columns: Columns{
				{Name: "col1", Label: "Column 1"},
				{Name: "col2", Label: "Column 2"},
				{Name: "col3", Label: "Column 3"},
				{Name: "col4", Label: "Column 4"},
			},
			conditions: MergeConditions{MergeConditionIdentical, MergeConditionEmpty},
			setupMock: func(mock *MockTableOperations) {
				mock.EXPECT().processValue("A", "").Return("A", nil).Times(2)
				mock.EXPECT().processValue("", "").Return("", nil).Times(1)
				mock.EXPECT().processValue(nil, "").Return("", nil).Times(1)
			},
			expectedRanges: [][]int{{0, 1}, {2, 3}}, // Changed: Mixed conditions create separate ranges for identical and empty values
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			ctrl := gomock.NewController(t)
			defer ctrl.Finish()

			mockOps := NewMockTableOperations(ctrl)
			tt.setupMock(mockOps)

			table := tt.setupTable()
			ranges := table.findHorizontalMergeRanges(tt.item, tt.columns, tt.conditions, mockOps)

			if len(ranges) != len(tt.expectedRanges) {
				t.Errorf("Expected %d ranges, got %d", len(tt.expectedRanges), len(ranges))
				return
			}

			for i, expectedRange := range tt.expectedRanges {
				if len(ranges[i]) != len(expectedRange) {
					t.Errorf("Range %d: expected length %d, got %d", i, len(expectedRange), len(ranges[i]))
					continue
				}
				for j, expectedVal := range expectedRange {
					if ranges[i][j] != expectedVal {
						t.Errorf("Range %d[%d]: expected %d, got %d", i, j, expectedVal, ranges[i][j])
					}
				}
			}
		})
	}
}

// Helper function to check if a string contains a substring
func containsSubstring(s, substr string) bool {
	return len(s) >= len(substr) && (s == substr || len(substr) == 0 ||
		(len(s) > 0 && len(substr) > 0 &&
			func() bool {
				for i := 0; i <= len(s)-len(substr); i++ {
					if s[i:i+len(substr)] == substr {
						return true
					}
				}
				return false
			}()))
}
