package spit

import (
	"reflect"
	"testing"
)

func TestTable_getDataStartRow(t *testing.T) {
	tests := []struct {
		name     string
		table    Table
		expected int
	}{
		{
			name: "No header, no columns",
			table: Table{
				WriteHeader: false,
				Columns:     Columns{},
			},
			expected: 1,
		},
		{
			name: "Header disabled",
			table: Table{
				WriteHeader: false,
				Columns: Columns{
					{Name: "col1", Label: "Column 1"},
					{Name: "col2", Label: "Column 2"},
				},
			},
			expected: 1,
		},
		{
			name: "Header enabled with simple columns",
			table: Table{
				WriteHeader: true,
				Columns: Columns{
					{Name: "col1", Label: "Column 1"},
					{Name: "col2", Label: "Column 2"},
				},
			},
			expected: 2,
		},
		{
			name: "Header enabled with nested columns",
			table: Table{
				WriteHeader: true,
				Columns: Columns{
					{
						Label: "Group 1",
						Columns: Columns{
							{Name: "col1", Label: "Column 1"},
							{Name: "col2", Label: "Column 2"},
						},
					},
				},
			},
			expected: 3,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := tt.table.getDataStartRow()
			if result != tt.expected {
				t.Errorf("getDataStartRow() = %v, want %v", result, tt.expected)
			}
		})
	}
}

func TestTable_getDataIndexFromRowIndex(t *testing.T) {
	tests := []struct {
		name     string
		table    Table
		rowIndex int
		expected int
	}{
		{
			name: "No header",
			table: Table{
				WriteHeader: false,
			},
			rowIndex: 5,
			expected: 5,
		},
		{
			name: "With header, simple columns",
			table: Table{
				WriteHeader: true,
				Columns: Columns{
					{Name: "col1", Label: "Column 1"},
				},
			},
			rowIndex: 3,
			expected: 1,
		},
		{
			name: "With header, nested columns",
			table: Table{
				WriteHeader: true,
				Columns: Columns{
					{
						Label: "Group",
						Columns: Columns{
							{Name: "col1", Label: "Column 1"},
						},
					},
				},
			},
			rowIndex: 4,
			expected: 1,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := tt.table.getDataIndexFromRowIndex(tt.rowIndex)
			if result != tt.expected {
				t.Errorf("getDataIndexFromRowIndex() = %v, want %v", result, tt.expected)
			}
		})
	}
}

func TestData_lookup(t *testing.T) {
	data := Data{
		"simple": "value1",
		"nested": Data{
			"level2": "value2",
			"deeper": Data{
				"level3": "value3",
			},
		},
		"empty": "",
	}

	tests := []struct {
		name     string
		keys     []string
		expected interface{}
		wantErr  bool
	}{
		{
			name:     "Simple key lookup",
			keys:     []string{"simple"},
			expected: "value1",
			wantErr:  false,
		},
		{
			name:     "Nested key lookup",
			keys:     []string{"nested", "level2"},
			expected: "value2",
			wantErr:  false,
		},
		{
			name:     "Deep nested key lookup",
			keys:     []string{"nested", "deeper", "level3"},
			expected: "value3",
			wantErr:  false,
		},
		{
			name:     "Key with spaces",
			keys:     []string{" simple "},
			expected: "value1",
			wantErr:  false,
		},
		{
			name:     "No keys provided",
			keys:     []string{},
			expected: nil,
			wantErr:  true,
		},
		{
			name:     "Key not found",
			keys:     []string{"nonexistent"},
			expected: nil,
			wantErr:  true,
		},
		{
			name:     "Malformed structure",
			keys:     []string{"simple", "invalid"},
			expected: nil,
			wantErr:  true,
		},
		{
			name:     "Nested key not found",
			keys:     []string{"nested", "nonexistent"},
			expected: nil,
			wantErr:  true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result, err := data.lookup(tt.keys...)
			if tt.wantErr {
				if err == nil {
					t.Errorf("lookup() expected error, got nil")
				}
			} else {
				if err != nil {
					t.Errorf("lookup() unexpected error: %v", err)
				}
				if result != tt.expected {
					t.Errorf("lookup() = %v, want %v", result, tt.expected)
				}
			}
		})
	}
}

func TestColumn_hasSubColumns(t *testing.T) {
	tests := []struct {
		name     string
		column   Column
		expected bool
	}{
		{
			name: "No sub-columns",
			column: Column{
				Name:  "simple",
				Label: "Simple Column",
			},
			expected: false,
		},
		{
			name: "Empty sub-columns",
			column: Column{
				Name:    "empty",
				Label:   "Empty Sub-columns",
				Columns: Columns{},
			},
			expected: false,
		},
		{
			name: "Has sub-columns",
			column: Column{
				Label: "Parent",
				Columns: Columns{
					{Name: "child1", Label: "Child 1"},
				},
			},
			expected: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := tt.column.hasSubColumns()
			if result != tt.expected {
				t.Errorf("hasSubColumns() = %v, want %v", result, tt.expected)
			}
		})
	}
}

func TestColumn_getColumnCount(t *testing.T) {
	tests := []struct {
		name     string
		column   Column
		expected int
	}{
		{
			name: "Leaf column",
			column: Column{
				Name:  "leaf",
				Label: "Leaf Column",
			},
			expected: 1,
		},
		{
			name: "Parent with one child",
			column: Column{
				Label: "Parent",
				Columns: Columns{
					{Name: "child1", Label: "Child 1"},
				},
			},
			expected: 1,
		},
		{
			name: "Parent with multiple children",
			column: Column{
				Label: "Parent",
				Columns: Columns{
					{Name: "child1", Label: "Child 1"},
					{Name: "child2", Label: "Child 2"},
					{Name: "child3", Label: "Child 3"},
				},
			},
			expected: 3,
		},
		{
			name: "Nested hierarchy",
			column: Column{
				Label: "Level1",
				Columns: Columns{
					{
						Label: "Level2A",
						Columns: Columns{
							{Name: "leaf1", Label: "Leaf 1"},
							{Name: "leaf2", Label: "Leaf 2"},
						},
					},
					{Name: "leaf3", Label: "Leaf 3"},
				},
			},
			expected: 3,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := tt.column.getColumnCount()
			if result != tt.expected {
				t.Errorf("getColumnCount() = %v, want %v", result, tt.expected)
			}
		})
	}
}

func TestColumns_getTotalColumnCount(t *testing.T) {
	tests := []struct {
		name     string
		columns  Columns
		expected int
	}{
		{
			name:     "Empty columns",
			columns:  Columns{},
			expected: 0,
		},
		{
			name: "Simple columns",
			columns: Columns{
				{Name: "col1", Label: "Column 1"},
				{Name: "col2", Label: "Column 2"},
			},
			expected: 2,
		},
		{
			name: "Mixed simple and nested",
			columns: Columns{
				{Name: "simple", Label: "Simple"},
				{
					Label: "Group",
					Columns: Columns{
						{Name: "nested1", Label: "Nested 1"},
						{Name: "nested2", Label: "Nested 2"},
					},
				},
			},
			expected: 3,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := tt.columns.getTotalColumnCount()
			if result != tt.expected {
				t.Errorf("getTotalColumnCount() = %v, want %v", result, tt.expected)
			}
		})
	}
}

func TestColumns_getFlattenedColumns(t *testing.T) {
	tests := []struct {
		name     string
		columns  Columns
		expected []Column
	}{
		{
			name:     "Empty columns",
			columns:  Columns{},
			expected: nil, // Use nil instead of []Column{} for empty slice comparison
		},
		{
			name: "Simple columns",
			columns: Columns{
				{Name: "col1", Label: "Column 1"},
				{Name: "col2", Label: "Column 2"},
			},
			expected: []Column{
				{Name: "col1", Label: "Column 1"},
				{Name: "col2", Label: "Column 2"},
			},
		},
		{
			name: "Nested columns",
			columns: Columns{
				{
					Label: "Group",
					Columns: Columns{
						{Name: "nested1", Label: "Nested 1"},
						{Name: "nested2", Label: "Nested 2"},
					},
				},
			},
			expected: []Column{
				{Name: "nested1", Label: "Nested 1"},
				{Name: "nested2", Label: "Nested 2"},
			},
		},
		{
			name: "Mixed simple and nested",
			columns: Columns{
				{Name: "simple", Label: "Simple"},
				{
					Label: "Group",
					Columns: Columns{
						{Name: "nested1", Label: "Nested 1"},
					},
				},
			},
			expected: []Column{
				{Name: "simple", Label: "Simple"},
				{Name: "nested1", Label: "Nested 1"},
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := tt.columns.getFlattenedColumns()
			if !reflect.DeepEqual(result, tt.expected) {
				t.Errorf("getFlattenedColumns() = %v, want %v", result, tt.expected)
			}
		})
	}
}

func TestColumns_getMaxDepth(t *testing.T) {
	tests := []struct {
		name     string
		columns  Columns
		expected int
	}{
		{
			name:     "Empty columns",
			columns:  Columns{},
			expected: 1,
		},
		{
			name: "Simple columns",
			columns: Columns{
				{Name: "col1", Label: "Column 1"},
				{Name: "col2", Label: "Column 2"},
			},
			expected: 1,
		},
		{
			name: "Two-level hierarchy",
			columns: Columns{
				{
					Label: "Group",
					Columns: Columns{
						{Name: "nested1", Label: "Nested 1"},
					},
				},
			},
			expected: 2,
		},
		{
			name: "Three-level hierarchy",
			columns: Columns{
				{
					Label: "Level1",
					Columns: Columns{
						{
							Label: "Level2",
							Columns: Columns{
								{Name: "leaf", Label: "Leaf"},
							},
						},
					},
				},
			},
			expected: 3,
		},
		{
			name: "Mixed depths",
			columns: Columns{
				{Name: "simple", Label: "Simple"},
				{
					Label: "Deep",
					Columns: Columns{
						{
							Label: "Deeper",
							Columns: Columns{
								{Name: "deepest", Label: "Deepest"},
							},
						},
					},
				},
			},
			expected: 3,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := tt.columns.getMaxDepth()
			if result != tt.expected {
				t.Errorf("getMaxDepth() = %v, want %v", result, tt.expected)
			}
		})
	}
}

func TestMergeConditions_anyMatch(t *testing.T) {
	tests := []struct {
		name       string
		conditions MergeConditions
		other      []MergeCondition
		expected   bool
	}{
		{
			name:       "Empty conditions",
			conditions: MergeConditions{},
			other:      []MergeCondition{MergeConditionIdentical},
			expected:   false,
		},
		{
			name:       "Empty other",
			conditions: MergeConditions{MergeConditionIdentical},
			other:      []MergeCondition{},
			expected:   false,
		},
		{
			name:       "Matching conditions",
			conditions: MergeConditions{MergeConditionIdentical, MergeConditionEmpty},
			other:      []MergeCondition{MergeConditionIdentical},
			expected:   true,
		},
		{
			name:       "No matching conditions",
			conditions: MergeConditions{MergeConditionIdentical},
			other:      []MergeCondition{MergeConditionEmpty},
			expected:   false,
		},
		{
			name:       "Multiple matches",
			conditions: MergeConditions{MergeConditionIdentical, MergeConditionEmpty},
			other:      []MergeCondition{MergeConditionEmpty, MergeConditionIdentical},
			expected:   true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := tt.conditions.anyMatch(tt.other)
			if result != tt.expected {
				t.Errorf("anyMatch() = %v, want %v", result, tt.expected)
			}
		})
	}
}

func TestMergeConditions_valuesShouldMerge(t *testing.T) {
	tests := []struct {
		name       string
		conditions MergeConditions
		value1     interface{}
		value2     interface{}
		expected   bool
	}{
		{
			name:       "No conditions",
			conditions: MergeConditions{},
			value1:     "test",
			value2:     "test",
			expected:   false,
		},
		{
			name:       "Identical values - should merge",
			conditions: MergeConditions{MergeConditionIdentical},
			value1:     "test",
			value2:     "test",
			expected:   true,
		},
		{
			name:       "Different values - should not merge",
			conditions: MergeConditions{MergeConditionIdentical},
			value1:     "test1",
			value2:     "test2",
			expected:   false,
		},
		{
			name:       "Empty values - should merge",
			conditions: MergeConditions{MergeConditionEmpty},
			value1:     "",
			value2:     "",
			expected:   true,
		},
		{
			name:       "Nil values - should merge",
			conditions: MergeConditions{MergeConditionEmpty},
			value1:     nil,
			value2:     nil,
			expected:   true,
		},
		{
			name:       "One empty, one nil - should merge",
			conditions: MergeConditions{MergeConditionEmpty},
			value1:     "",
			value2:     nil,
			expected:   true,
		},
		{
			name:       "Identical empty values shouldn't merge with identical condition",
			conditions: MergeConditions{MergeConditionIdentical},
			value1:     "",
			value2:     "",
			expected:   false,
		},
		{
			name:       "Non-empty values shouldn't merge with empty condition",
			conditions: MergeConditions{MergeConditionEmpty},
			value1:     "test",
			value2:     "test",
			expected:   false,
		},
		{
			name:       "Multiple conditions - identical match",
			conditions: MergeConditions{MergeConditionIdentical, MergeConditionEmpty},
			value1:     "test",
			value2:     "test",
			expected:   true,
		},
		{
			name:       "Multiple conditions - empty match",
			conditions: MergeConditions{MergeConditionIdentical, MergeConditionEmpty},
			value1:     "",
			value2:     nil,
			expected:   true,
		},
		{
			name:       "Whitespace handling",
			conditions: MergeConditions{MergeConditionIdentical},
			value1:     " test ",
			value2:     " test ",
			expected:   true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := tt.conditions.valuesShouldMerge(tt.value1, tt.value2)
			if result != tt.expected {
				t.Errorf("valuesShouldMerge() = %v, want %v", result, tt.expected)
			}
		})
	}
}

func TestBorders_hasBorders(t *testing.T) {
	tests := []struct {
		name     string
		borders  *Borders
		expected bool
	}{
		{
			name:     "Nil borders",
			borders:  nil,
			expected: false,
		},
		{
			name:     "No borders",
			borders:  &Borders{},
			expected: false,
		},
		{
			name: "All borders none",
			borders: &Borders{
				Left:   &Border{Style: BorderStyleNone},
				Right:  &Border{Style: BorderStyleNone},
				Top:    &Border{Style: BorderStyleNone},
				Bottom: &Border{Style: BorderStyleNone},
			},
			expected: false,
		},
		{
			name: "Left border only",
			borders: &Borders{
				Left: &Border{Style: BorderStyleThin},
			},
			expected: true,
		},
		{
			name: "Right border only",
			borders: &Borders{
				Right: &Border{Style: BorderStyleMedium},
			},
			expected: true,
		},
		{
			name: "Top border only",
			borders: &Borders{
				Top: &Border{Style: BorderStyleThick},
			},
			expected: true,
		},
		{
			name: "Bottom border only",
			borders: &Borders{
				Bottom: &Border{Style: BorderStyleDashed},
			},
			expected: true,
		},
		{
			name: "Multiple borders",
			borders: &Borders{
				Left:  &Border{Style: BorderStyleThin},
				Right: &Border{Style: BorderStyleThin},
			},
			expected: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			var result bool
			if tt.borders != nil {
				result = tt.borders.hasBorders()
			}
			if result != tt.expected {
				t.Errorf("hasBorders() = %v, want %v", result, tt.expected)
			}
		})
	}
}

func TestBorders_SetInner(t *testing.T) {
	borders := &Borders{}
	result := borders.SetInner(BorderStyleThin)

	if result != borders {
		t.Errorf("SetInner() should return the same borders instance")
	}

	if borders.Inner == nil {
		t.Errorf("SetInner() should create Inner borders")
		return
	}

	expectedBorder := &Border{Style: BorderStyleThin}
	if !reflect.DeepEqual(borders.Inner.Left, expectedBorder) {
		t.Errorf("SetInner() Left border = %v, want %v", borders.Inner.Left, expectedBorder)
	}
	if !reflect.DeepEqual(borders.Inner.Right, expectedBorder) {
		t.Errorf("SetInner() Right border = %v, want %v", borders.Inner.Right, expectedBorder)
	}
	if !reflect.DeepEqual(borders.Inner.Top, expectedBorder) {
		t.Errorf("SetInner() Top border = %v, want %v", borders.Inner.Top, expectedBorder)
	}
	if !reflect.DeepEqual(borders.Inner.Bottom, expectedBorder) {
		t.Errorf("SetInner() Bottom border = %v, want %v", borders.Inner.Bottom, expectedBorder)
	}
}

func TestNewBorderOptions(t *testing.T) {
	tests := []struct {
		name  string
		style BorderStyle
	}{
		{"Thin border", BorderStyleThin},
		{"Medium border", BorderStyleMedium},
		{"Thick border", BorderStyleThick},
		{"Dashed border", BorderStyleDashed},
		{"Dotted border", BorderStyleDotted},
		{"Double border", BorderStyleDouble},
		{"None border", BorderStyleNone},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := NewBorderOptions(tt.style)

			expectedBorder := &Border{Style: tt.style}
			if !reflect.DeepEqual(result.Left, expectedBorder) {
				t.Errorf("NewBorderOptions() Left = %v, want %v", result.Left, expectedBorder)
			}
			if !reflect.DeepEqual(result.Right, expectedBorder) {
				t.Errorf("NewBorderOptions() Right = %v, want %v", result.Right, expectedBorder)
			}
			if !reflect.DeepEqual(result.Top, expectedBorder) {
				t.Errorf("NewBorderOptions() Top = %v, want %v", result.Top, expectedBorder)
			}
			if !reflect.DeepEqual(result.Bottom, expectedBorder) {
				t.Errorf("NewBorderOptions() Bottom = %v, want %v", result.Bottom, expectedBorder)
			}
		})
	}
}

func TestAlignment_getAlignmentValues(t *testing.T) {
	tests := []struct {
		name               string
		alignment          Alignment
		expectedHorizontal string
		expectedVertical   string
	}{
		{"Left", AlignmentLeft, "left", "top"},
		{"Center", AlignmentCenter, "center", "top"},
		{"Right", AlignmentRight, "right", "top"},
		{"Top", AlignmentTop, "left", "top"},
		{"Middle", AlignmentMiddle, "left", "center"},
		{"Bottom", AlignmentBottom, "left", "bottom"},
		{"Center Middle", AlignmentCenterMiddle, "center", "center"},
		{"Left Middle", AlignmentLeftMiddle, "left", "center"},
		{"Right Middle", AlignmentRightMiddle, "right", "center"},
		{"None", AlignmentNone, "left", "top"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			horizontal, vertical := tt.alignment.getAlignmentValues()
			if horizontal != tt.expectedHorizontal {
				t.Errorf("getAlignmentValues() horizontal = %v, want %v", horizontal, tt.expectedHorizontal)
			}
			if vertical != tt.expectedVertical {
				t.Errorf("getAlignmentValues() vertical = %v, want %v", vertical, tt.expectedVertical)
			}
		})
	}
}
