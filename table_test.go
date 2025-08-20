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
			result := tt.table.GetDataStartRow()
			if result != tt.expected {
				t.Errorf("GetDataStartRow() = %v, want %v", result, tt.expected)
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
			result := tt.table.GetDataIndexFromRowIndex(tt.rowIndex)
			if result != tt.expected {
				t.Errorf("GetDataIndexFromRowIndex() = %v, want %v", result, tt.expected)
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
		name      string
		keys      []string
		expected  interface{}
		wantErr   bool
		wantFound bool
	}{
		{
			name:      "Simple key Lookup",
			keys:      []string{"simple"},
			expected:  "value1",
			wantErr:   false,
			wantFound: true,
		},
		{
			name:      "Nested key Lookup",
			keys:      []string{"nested", "level2"},
			expected:  "value2",
			wantErr:   false,
			wantFound: true,
		},
		{
			name:      "Deep nested key Lookup",
			keys:      []string{"nested", "deeper", "level3"},
			expected:  "value3",
			wantFound: true,
		},
		{
			name:      "Key with spaces",
			keys:      []string{" simple "},
			expected:  "value1",
			wantErr:   false,
			wantFound: true,
		},
		{
			name:      "No keys provided",
			keys:      []string{},
			expected:  nil,
			wantErr:   true,
			wantFound: false,
		},
		{
			name:      "Key not found",
			keys:      []string{"nonexistent"},
			expected:  nil,
			wantErr:   false,
			wantFound: false,
		},
		{
			name:      "Malformed structure",
			keys:      []string{"simple", "invalid"},
			expected:  nil,
			wantErr:   true,
			wantFound: false,
		},
		{
			name:      "Nested key not found",
			keys:      []string{"nested", "nonexistent"},
			expected:  nil,
			wantErr:   false,
			wantFound: false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result, err, found := data.Lookup(tt.keys...)
			if found != tt.wantFound {
				if found != tt.wantFound {
					t.Errorf("Lookup() found = %v, want %v (keys: %v)", found, tt.wantFound, tt.keys)
				}
			}
			if tt.wantErr {
				if err == nil {
					t.Errorf("Lookup() expected error, got nil")
				}
			} else {
				if err != nil {
					t.Errorf("Lookup() unexpected error: %v", err)
				}
				if result != tt.expected {
					t.Errorf("Lookup() = %v, want %v", result, tt.expected)
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
			result := tt.column.HasSubColumns()
			if result != tt.expected {
				t.Errorf("HasSubColumns() = %v, want %v", result, tt.expected)
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
			result := tt.column.CountSubColumns()
			if result != tt.expected {
				t.Errorf("CountSubColumns() = %v, want %v", result, tt.expected)
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
			result := tt.columns.GetTotalColumnCount()
			if result != tt.expected {
				t.Errorf("GetTotalColumnCount() = %v, want %v", result, tt.expected)
			}
		})
	}
}

func TestColumns_getFlattenedColumns(t *testing.T) {
	tests := []struct {
		name     string
		columns  Columns
		expected Columns
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
			expected: Columns{
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
			expected: Columns{
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
			expected: Columns{
				{Name: "simple", Label: "Simple"},
				{Name: "nested1", Label: "Nested 1"},
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := tt.columns.GetFlattenedColumns()
			if !reflect.DeepEqual(result, tt.expected) {
				t.Errorf("GetFlattenedColumns() = %v, want %v", result, tt.expected)
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
			result := tt.columns.GetMaxDepth()
			if result != tt.expected {
				t.Errorf("GetMaxDepth() = %v, want %v", result, tt.expected)
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
			result := tt.conditions.AnyMatch(tt.other)
			if result != tt.expected {
				t.Errorf("AnyMatch() = %v, want %v", result, tt.expected)
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
			result := tt.conditions.ValuesShouldMerge(tt.value1, tt.value2)
			if result != tt.expected {
				t.Errorf("ValuesShouldMerge() = %v, want %v", result, tt.expected)
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
				result = tt.borders.HasBorders()
			}
			if result != tt.expected {
				t.Errorf("HasBorders() = %v, want %v", result, tt.expected)
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

func TestNewTable(t *testing.T) {
	data := DataSlice{
		{"name": "John", "age": 30},
		{"name": "Jane", "age": 25},
	}
	columns := Columns{
		NewColumn("name", "Name"),
		NewColumn("age", "Age"),
	}

	table := NewTable(data, columns, true)

	if table.Data == nil {
		t.Errorf("NewTable() Data should not be nil")
	}
	if len(table.Data) != 2 {
		t.Errorf("NewTable() Data length = %v, want %v", len(table.Data), 2)
	}
	if table.Columns == nil {
		t.Errorf("NewTable() Columns should not be nil")
	}
	if len(table.Columns) != 2 {
		t.Errorf("NewTable() Columns length = %v, want %v", len(table.Columns), 2)
	}
	if !table.WriteHeader {
		t.Errorf("NewTable() WriteHeader = %v, want %v", table.WriteHeader, true)
	}
}

func TestTable_WithRowOptions(t *testing.T) {
	table := &Table{}
	rowOptions := RowOptionsMap{
		0: *NewRowOptions(0),
		1: *NewRowOptions(1),
	}

	result := table.WithRowOptions(rowOptions)

	if result != table {
		t.Errorf("WithRowOptions() should return the same table instance")
	}
	if table.RowOptionsMap == nil {
		t.Errorf("WithRowOptions() should set RowOptionsMap")
	}
	if len(table.RowOptionsMap) != 2 {
		t.Errorf("WithRowOptions() RowOptionsMap length = %v, want %v", len(table.RowOptionsMap), 2)
	}
}

func TestTable_WithCellOptions(t *testing.T) {
	table := &Table{}
	cellOptions := CellOptionsMap{
		0: {
			0: *NewCellOptions(0, 0),
			1: *NewCellOptions(1, 0),
		},
	}

	result := table.WithCellOptions(cellOptions)

	if result != table {
		t.Errorf("WithCellOptions() should return the same table instance")
	}
	if table.CellOptionsMap == nil {
		t.Errorf("WithCellOptions() should set CellOptionsMap")
	}
	if len(table.CellOptionsMap) != 1 {
		t.Errorf("WithCellOptions() CellOptionsMap length = %v, want %v", len(table.CellOptionsMap), 1)
	}
}

func TestNewColumn(t *testing.T) {
	name := "test_name"
	label := "Test Label"

	column := NewColumn(name, label)

	if column.Name != name {
		t.Errorf("NewColumn() Name = %v, want %v", column.Name, name)
	}
	if column.Label != label {
		t.Errorf("NewColumn() Label = %v, want %v", column.Label, label)
	}
}

func TestColumn_WithFormat(t *testing.T) {
	column := &Column{}
	format := "2006-01-02"

	result := column.WithFormat(format)

	if result != column {
		t.Errorf("WithFormat() should return the same column instance")
	}
	if column.Format != format {
		t.Errorf("WithFormat() Format = %v, want %v", column.Format, format)
	}
}

func TestColumn_WithMerge(t *testing.T) {
	column := &Column{}
	merge := NewMergeRules(MergeConditions{MergeConditionIdentical}, MergeConditions{MergeConditionEmpty})

	result := column.WithMerge(merge)

	if result != column {
		t.Errorf("WithMerge() should return the same column instance")
	}
	if column.Merge != merge {
		t.Errorf("WithMerge() Merge should be set")
	}
}

func TestColumn_WithBorders(t *testing.T) {
	column := &Column{}
	borders := NewBordersBoundaries(BorderStyleThin)

	result := column.WithBorders(borders)

	if result != column {
		t.Errorf("WithBorders() should return the same column instance")
	}
	if column.Borders != borders {
		t.Errorf("WithBorders() Borders should be set")
	}
}

func TestColumn_WithStyle(t *testing.T) {
	column := &Column{}
	style := &Style{Bold: true, FontSize: 12}

	result := column.WithStyle(style)

	if result != column {
		t.Errorf("WithStyle() should return the same column instance")
	}
	if column.Style != style {
		t.Errorf("WithStyle() Style should be set")
	}
}

func TestColumn_WithSubColumns(t *testing.T) {
	column := &Column{}
	subColumns := Columns{
		NewColumn("sub1", "Sub 1"),
		NewColumn("sub2", "Sub 2"),
	}

	result := column.WithSubColumns(subColumns)

	if result != column {
		t.Errorf("WithSubColumns() should return the same column instance")
	}
	if len(column.Columns) != 2 {
		t.Errorf("WithSubColumns() Columns length = %v, want %v", len(column.Columns), 2)
	}
}

func TestColumn_AddSubColumn(t *testing.T) {
	tests := []struct {
		name           string
		initialColumns Columns
		newColumn      *Column
		expectedCount  int
	}{
		{
			name:           "Add to empty columns",
			initialColumns: nil,
			newColumn:      NewColumn("new", "New"),
			expectedCount:  1,
		},
		{
			name: "Add to existing columns",
			initialColumns: Columns{
				NewColumn("existing", "Existing"),
			},
			newColumn:     NewColumn("new", "New"),
			expectedCount: 2,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			column := &Column{Columns: tt.initialColumns}

			result := column.AddSubColumn(tt.newColumn)

			if result != column {
				t.Errorf("AddSubColumn() should return the same column instance")
			}
			if len(column.Columns) != tt.expectedCount {
				t.Errorf("AddSubColumn() Columns length = %v, want %v", len(column.Columns), tt.expectedCount)
			}
			if column.Columns[len(column.Columns)-1] != tt.newColumn {
				t.Errorf("AddSubColumn() last column should be the new column")
			}
		})
	}
}

func TestColumn_RemoveSubColumn(t *testing.T) {
	tests := []struct {
		name           string
		initialColumns Columns
		removeByName   string
		expectedCount  int
	}{
		{
			name:           "Remove from empty columns",
			initialColumns: Columns{},
			removeByName:   "nonexistent",
			expectedCount:  0,
		},
		{
			name: "Remove existing column",
			initialColumns: Columns{
				NewColumn("keep", "Keep"),
				NewColumn("remove", "Remove"),
				NewColumn("keep2", "Keep2"),
			},
			removeByName:  "remove",
			expectedCount: 2,
		},
		{
			name: "Remove nonexistent column",
			initialColumns: Columns{
				NewColumn("keep", "Keep"),
			},
			removeByName:  "nonexistent",
			expectedCount: 1,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			column := &Column{Columns: tt.initialColumns}

			result := column.RemoveSubColumn(tt.removeByName)

			if result != column {
				t.Errorf("RemoveSubColumn() should return the same column instance")
			}
			if len(column.Columns) != tt.expectedCount {
				t.Errorf("RemoveSubColumn() Columns length = %v, want %v", len(column.Columns), tt.expectedCount)
			}
			// Check that the removed column is not present
			for _, subCol := range column.Columns {
				if subCol.Name == tt.removeByName {
					t.Errorf("RemoveSubColumn() should have removed column with name %v", tt.removeByName)
				}
			}
		})
	}
}

func TestNewRowOptions(t *testing.T) {
	rowIndex := 5

	rowOptions := NewRowOptions(rowIndex)

	if rowOptions.RowIndex != rowIndex {
		t.Errorf("NewRowOptions() RowIndex = %v, want %v", rowOptions.RowIndex, rowIndex)
	}
}

func TestRowOptions_WithBorder(t *testing.T) {
	rowOptions := &RowOptions{}
	borders := NewBordersBoundaries(BorderStyleThin)

	result := rowOptions.WithBorder(borders)

	if result != rowOptions {
		t.Errorf("WithBorder() should return the same RowOptions instance")
	}
	if rowOptions.Border != borders {
		t.Errorf("WithBorder() Border should be set")
	}
}

func TestRowOptions_WithStyle(t *testing.T) {
	rowOptions := &RowOptions{}
	style := &Style{Bold: true}

	result := rowOptions.WithStyle(style)

	if result != rowOptions {
		t.Errorf("WithStyle() should return the same RowOptions instance")
	}
	if rowOptions.Style != style {
		t.Errorf("WithStyle() Style should be set")
	}
}

func TestRowOptions_WithMerge(t *testing.T) {
	rowOptions := &RowOptions{}
	merge := NewMergeRules(MergeConditions{MergeConditionIdentical}, MergeConditions{})

	result := rowOptions.WithMerge(merge)

	if result != rowOptions {
		t.Errorf("WithMerge() should return the same RowOptions instance")
	}
	if rowOptions.Merge != merge {
		t.Errorf("WithMerge() Merge should be set")
	}
}

func TestRowOptions_WithMergeable(t *testing.T) {
	rowOptions := &RowOptions{}

	result := rowOptions.WithMergeable(true)

	if result != rowOptions {
		t.Errorf("WithMergeable() should return the same RowOptions instance")
	}
	if !rowOptions.Mergeable {
		t.Errorf("WithMergeable() Mergeable should be true")
	}
}

func TestNewCellOptions(t *testing.T) {
	rowIndex := 3
	colIndex := 7

	cellOptions := NewCellOptions(rowIndex, colIndex)

	if cellOptions.RowIndex != rowIndex {
		t.Errorf("NewCellOptions() RowIndex = %v, want %v", cellOptions.RowIndex, rowIndex)
	}
	if cellOptions.ColIndex != colIndex {
		t.Errorf("NewCellOptions() ColIndex = %v, want %v", cellOptions.ColIndex, colIndex)
	}
}

func TestCellOptions_WithBorder(t *testing.T) {
	cellOptions := &CellOptions{}
	borders := NewBordersBoundaries(BorderStyleMedium)

	result := cellOptions.WithBorder(borders)

	if result != cellOptions {
		t.Errorf("WithBorder() should return the same CellOptions instance")
	}
	if cellOptions.Border != borders {
		t.Errorf("WithBorder() Border should be set")
	}
}

func TestCellOptions_WithStyle(t *testing.T) {
	cellOptions := &CellOptions{}
	style := &Style{Italic: true}

	result := cellOptions.WithStyle(style)

	if result != cellOptions {
		t.Errorf("WithStyle() should return the same CellOptions instance")
	}
	if cellOptions.Style != style {
		t.Errorf("WithStyle() Style should be set")
	}
}

func TestCellOptions_WithMergeable(t *testing.T) {
	cellOptions := &CellOptions{}

	result := cellOptions.WithMergeable(false)

	if result != cellOptions {
		t.Errorf("WithMergeable() should return the same CellOptions instance")
	}
	if cellOptions.Mergeable {
		t.Errorf("WithMergeable() Mergeable should be false")
	}
}

func TestNewMergeRules(t *testing.T) {
	vertical := MergeConditions{MergeConditionIdentical}
	horizontal := MergeConditions{MergeConditionEmpty}

	mergeRules := NewMergeRules(vertical, horizontal)

	if !reflect.DeepEqual(mergeRules.Vertical, vertical) {
		t.Errorf("NewMergeRules() Vertical = %v, want %v", mergeRules.Vertical, vertical)
	}
	if !reflect.DeepEqual(mergeRules.Horizontal, horizontal) {
		t.Errorf("NewMergeRules() Horizontal = %v, want %v", mergeRules.Horizontal, horizontal)
	}
}

func TestNewBorder(t *testing.T) {
	style := BorderStyleThick

	border := NewBorder(style)

	if border.Style != style {
		t.Errorf("NewBorder() Style = %v, want %v", border.Style, style)
	}
}

func TestNewBorders(t *testing.T) {
	left := BorderStyleThin
	right := BorderStyleMedium
	top := BorderStyleThick
	bottom := BorderStyleDashed

	borders := NewBorders(left, right, top, bottom)

	if borders.Left == nil || borders.Left.Style != left {
		t.Errorf("NewBorders() Left style = %v, want %v", borders.Left.Style, left)
	}
	if borders.Right == nil || borders.Right.Style != right {
		t.Errorf("NewBorders() Right style = %v, want %v", borders.Right.Style, right)
	}
	if borders.Top == nil || borders.Top.Style != top {
		t.Errorf("NewBorders() Top style = %v, want %v", borders.Top.Style, top)
	}
	if borders.Bottom == nil || borders.Bottom.Style != bottom {
		t.Errorf("NewBorders() Bottom style = %v, want %v", borders.Bottom.Style, bottom)
	}
}

func TestNewBordersBoundaries(t *testing.T) {
	style := BorderStyleDouble

	borders := NewBordersBoundaries(style)

	expectedBorder := &Border{Style: style}
	if !reflect.DeepEqual(borders.Left, expectedBorder) {
		t.Errorf("NewBordersBoundaries() Left = %v, want %v", borders.Left, expectedBorder)
	}
	if !reflect.DeepEqual(borders.Right, expectedBorder) {
		t.Errorf("NewBordersBoundaries() Right = %v, want %v", borders.Right, expectedBorder)
	}
	if !reflect.DeepEqual(borders.Top, expectedBorder) {
		t.Errorf("NewBordersBoundaries() Top = %v, want %v", borders.Top, expectedBorder)
	}
	if !reflect.DeepEqual(borders.Bottom, expectedBorder) {
		t.Errorf("NewBordersBoundaries() Bottom = %v, want %v", borders.Bottom, expectedBorder)
	}
}

func TestBorders_SetBoundaries(t *testing.T) {
	borders := &Borders{}
	style := BorderStyleThin

	result := borders.SetBoundaries(style)

	if result != borders {
		t.Errorf("SetBoundaries() should return the same borders instance")
	}
	if borders.Left == nil || borders.Left.Style != style {
		t.Errorf("SetBoundaries() Left style = %v, want %v", borders.Left.Style, style)
	}
	if borders.Right == nil || borders.Right.Style != style {
		t.Errorf("SetBoundaries() Right style = %v, want %v", borders.Right.Style, style)
	}
	if borders.Top == nil || borders.Top.Style != style {
		t.Errorf("SetBoundaries() Top style = %v, want %v", borders.Top.Style, style)
	}
	if borders.Bottom == nil || borders.Bottom.Style != style {
		t.Errorf("SetBoundaries() Bottom style = %v, want %v", borders.Bottom.Style, style)
	}
}

func TestBorders_SetVertical(t *testing.T) {
	borders := &Borders{}
	style := BorderStyleMedium

	result := borders.SetVertical(style)

	if result != borders {
		t.Errorf("SetVertical() should return the same borders instance")
	}
	if borders.Left == nil || borders.Left.Style != style {
		t.Errorf("SetVertical() Left style = %v, want %v", borders.Left.Style, style)
	}
	if borders.Right == nil || borders.Right.Style != style {
		t.Errorf("SetVertical() Right style = %v, want %v", borders.Right.Style, style)
	}
}

func TestBorders_SetHorizontal(t *testing.T) {
	borders := &Borders{}
	style := BorderStyleThick

	result := borders.SetHorizontal(style)

	if result != borders {
		t.Errorf("SetHorizontal() should return the same borders instance")
	}
	if borders.Top == nil || borders.Top.Style != style {
		t.Errorf("SetHorizontal() Top style = %v, want %v", borders.Top.Style, style)
	}
	if borders.Bottom == nil || borders.Bottom.Style != style {
		t.Errorf("SetHorizontal() Bottom style = %v, want %v", borders.Bottom.Style, style)
	}
}

func TestBorders_SetLeft(t *testing.T) {
	tests := []struct {
		name           string
		initialBorders *Borders
		style          BorderStyle
	}{
		{
			name:           "Set left on new borders",
			initialBorders: &Borders{},
			style:          BorderStyleThin,
		},
		{
			name: "Update existing left border",
			initialBorders: &Borders{
				Left: NewBorder(BorderStyleMedium),
			},
			style: BorderStyleThick,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := tt.initialBorders.SetLeft(tt.style)

			if result != tt.initialBorders {
				t.Errorf("SetLeft() should return the same borders instance")
			}
			if tt.initialBorders.Left == nil || tt.initialBorders.Left.Style != tt.style {
				t.Errorf("SetLeft() Left style = %v, want %v", tt.initialBorders.Left.Style, tt.style)
			}
		})
	}
}

func TestBorders_SetRight(t *testing.T) {
	tests := []struct {
		name           string
		initialBorders *Borders
		style          BorderStyle
	}{
		{
			name:           "Set right on new borders",
			initialBorders: &Borders{},
			style:          BorderStyleDashed,
		},
		{
			name: "Update existing right border",
			initialBorders: &Borders{
				Right: NewBorder(BorderStyleThin),
			},
			style: BorderStyleDouble,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := tt.initialBorders.SetRight(tt.style)

			if result != tt.initialBorders {
				t.Errorf("SetRight() should return the same borders instance")
			}
			if tt.initialBorders.Right == nil || tt.initialBorders.Right.Style != tt.style {
				t.Errorf("SetRight() Right style = %v, want %v", tt.initialBorders.Right.Style, tt.style)
			}
		})
	}
}

func TestBorders_SetTop(t *testing.T) {
	tests := []struct {
		name           string
		initialBorders *Borders
		style          BorderStyle
	}{
		{
			name:           "Set top on new borders",
			initialBorders: &Borders{},
			style:          BorderStyleDotted,
		},
		{
			name: "Update existing top border",
			initialBorders: &Borders{
				Top: NewBorder(BorderStyleThin),
			},
			style: BorderStyleThick,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := tt.initialBorders.SetTop(tt.style)

			if result != tt.initialBorders {
				t.Errorf("SetTop() should return the same borders instance")
			}
			if tt.initialBorders.Top == nil || tt.initialBorders.Top.Style != tt.style {
				t.Errorf("SetTop() Top style = %v, want %v", tt.initialBorders.Top.Style, tt.style)
			}
		})
	}
}

func TestBorders_SetBottom(t *testing.T) {
	tests := []struct {
		name           string
		initialBorders *Borders
		style          BorderStyle
	}{
		{
			name:           "Set bottom on new borders",
			initialBorders: &Borders{},
			style:          BorderStyleMedium,
		},
		{
			name: "Update existing bottom border",
			initialBorders: &Borders{
				Bottom: NewBorder(BorderStyleDashed),
			},
			style: BorderStyleNone,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := tt.initialBorders.SetBottom(tt.style)

			if result != tt.initialBorders {
				t.Errorf("SetBottom() should return the same borders instance")
			}
			if tt.initialBorders.Bottom == nil || tt.initialBorders.Bottom.Style != tt.style {
				t.Errorf("SetBottom() Bottom style = %v, want %v", tt.initialBorders.Bottom.Style, tt.style)
			}
		})
	}
}

func TestAlignment_GetAlignmentValues(t *testing.T) {
	tests := []struct {
		name               string
		alignment          Alignment
		expectedHorizontal string
		expectedVertical   string
	}{
		{
			name:               "AlignmentNone",
			alignment:          AlignmentNone,
			expectedHorizontal: "left",
			expectedVertical:   "top",
		},
		{
			name:               "AlignmentLeft",
			alignment:          AlignmentLeft,
			expectedHorizontal: "left",
			expectedVertical:   "top",
		},
		{
			name:               "AlignmentCenter",
			alignment:          AlignmentCenter,
			expectedHorizontal: "center",
			expectedVertical:   "top",
		},
		{
			name:               "AlignmentRight",
			alignment:          AlignmentRight,
			expectedHorizontal: "right",
			expectedVertical:   "top",
		},
		{
			name:               "AlignmentTop",
			alignment:          AlignmentTop,
			expectedHorizontal: "left",
			expectedVertical:   "top",
		},
		{
			name:               "AlignmentMiddle",
			alignment:          AlignmentMiddle,
			expectedHorizontal: "left",
			expectedVertical:   "center",
		},
		{
			name:               "AlignmentBottom",
			alignment:          AlignmentBottom,
			expectedHorizontal: "left",
			expectedVertical:   "bottom",
		},
		{
			name:               "AlignmentCenterMiddle",
			alignment:          AlignmentCenterMiddle,
			expectedHorizontal: "center",
			expectedVertical:   "center",
		},
		{
			name:               "AlignmentLeftMiddle",
			alignment:          AlignmentLeftMiddle,
			expectedHorizontal: "left",
			expectedVertical:   "center",
		},
		{
			name:               "AlignmentRightMiddle",
			alignment:          AlignmentRightMiddle,
			expectedHorizontal: "right",
			expectedVertical:   "center",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			horizontal, vertical := tt.alignment.GetAlignmentValues()
			if horizontal != tt.expectedHorizontal {
				t.Errorf("GetAlignmentValues() horizontal = %v, want %v", horizontal, tt.expectedHorizontal)
			}
			if vertical != tt.expectedVertical {
				t.Errorf("GetAlignmentValues() vertical = %v, want %v", vertical, tt.expectedVertical)
			}
		})
	}
}
