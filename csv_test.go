package spit

import (
	"bytes"
	stdcsv "encoding/csv"
	"os"
	"strings"
	"testing"
	"time"
)

// Common test data
var (
	testData = DataSlice{
		{"name": "John", "age": 30, "city": "New York"},
		{"name": "Jane", "age": 25, "city": "Los Angeles"},
	}

	testTable = &Table{
		Data: testData,
		Columns: Columns{
			{Name: "name", Label: "Name"},
			{Name: "age", Label: "Age"},
			{Name: "city", Label: "City"},
		},
		WriteHeader: true,
	}
)

// TestExportCSV tests the main ExportCSV function
func TestExportCSV(t *testing.T) {
	tests := []struct {
		name      string
		separator string
		params    FileWriteParams
		table     *Table
		wantExt   string
		contains  string
	}{
		{
			name:      "DefaultParameters",
			separator: ",",
			params:    FileWriteParams{UseTempFile: true, OverwriteFile: true},
			table:     testTable,
			wantExt:   ".csv",
			contains:  "Name,Age,City",
		},
		{
			name:      "CustomSeparator",
			separator: ";",
			params:    FileWriteParams{UseTempFile: true, OverwriteFile: true},
			table:     testTable,
			wantExt:   ".csv",
			contains:  "Name;Age;City",
		},
		{
			name:      "WithListSeparator",
			separator: ",",
			params:    FileWriteParams{UseTempFile: true, OverwriteFile: true},
			table: &Table{
				Data: DataSlice{{"name": "John", "tags": []interface{}{"developer", "go", "testing"}}},
				Columns: Columns{
					{Name: "name", Label: "Name"},
					{Name: "tags", Label: "Tags"},
				},
				WriteHeader:   true,
				ListSeparator: "|",
			},
			wantExt:  ".csv",
			contains: "developer|go|testing",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			// Create temp file that will be auto-cleaned
			tempFile, err := os.CreateTemp("", "csv_test_*.csv")
			if err != nil {
				t.Fatalf("Failed to create temp file: %v", err)
			}
			defer func(name string) {
				_ = os.Remove(name)
			}(tempFile.Name())
			_ = tempFile.Close()

			tt.params.Filename = strings.TrimSuffix(tempFile.Name(), ".csv")

			result, err := ExportCSV(tt.separator, tt.table, tt.params)
			if err != nil {
				t.Errorf("ExportCSV should not return error, got: %v", err)
				return
			}
			defer func(name string) {
				_ = os.Remove(name)
			}(result.Filepath)

			if !strings.HasSuffix(result.Filepath, tt.wantExt) {
				t.Errorf("Expected Extension %s, got %s", tt.wantExt, result.Filepath)
			}

			if tt.contains != "" {
				content, err := os.ReadFile(result.Filepath)
				if err != nil {
					t.Errorf("Failed to read CSV file: %v", err)
					return
				}
				if !strings.Contains(string(content), tt.contains) {
					t.Errorf("CSV should contain '%s'", tt.contains)
				}
			}
		})
	}
}

// TestCSV_writeData tests the writeData method
func TestCSV_writeData(t *testing.T) {
	createCSVInstance := func(table *Table, separator string) (*csv, *bytes.Buffer) {
		buf := &bytes.Buffer{}
		return &csv{
			writer:    stdcsv.NewWriter(buf),
			separator: separator,
			table:     table,
		}, buf
	}

	t.Run("BasicData", func(t *testing.T) {
		csvInstance, buf := createCSVInstance(testTable, ",")

		if err := csvInstance.writeData(); err != nil {
			t.Errorf("writeData should not return error, got: %v", err)
		}

		lines := strings.Split(strings.TrimSpace(buf.String()), "\n")
		expected := []string{"Name,Age,City", "John,30,New York", "Jane,25,Los Angeles"}

		for i, exp := range expected {
			if i >= len(lines) || lines[i] != exp {
				t.Errorf("Line %d: expected '%s', got '%s'", i, exp, lines[i])
			}
		}
	})

	t.Run("WithoutHeaders", func(t *testing.T) {
		table := *testTable // Copy
		table.WriteHeader = false
		csvInstance, buf := createCSVInstance(&table, ",")

		if err := csvInstance.writeData(); err != nil {
			t.Errorf("writeData should not return error, got: %v", err)
		}

		lines := strings.Split(strings.TrimSpace(buf.String()), "\n")
		if len(lines) != 2 {
			t.Errorf("Expected 2 lines (data only), got %d", len(lines))
		}
	})

	t.Run("MissingData", func(t *testing.T) {
		table := &Table{
			Data: DataSlice{{"name": "John"}}, // Missing age field
			Columns: Columns{
				{Name: "name", Label: "Name"},
				{Name: "age", Label: "Age"}, // Missing in data
			},
		}
		csvInstance, _ := createCSVInstance(table, ",")

		err := csvInstance.writeData()
		if err != nil {
			t.Errorf("Unexpected error: %v", err)
		}
	})

	t.Run("WithFormats", func(t *testing.T) {
		table := &Table{
			Data: DataSlice{{"name": "John", "date": time.Date(2023, 1, 15, 0, 0, 0, 0, time.UTC)}},
			Columns: Columns{
				{Name: "name", Label: "Name"},
				{Name: "date", Label: "Date", Format: "2006-01-02"},
			},
		}
		csvInstance, buf := createCSVInstance(table, ",")

		if err := csvInstance.writeData(); err != nil {
			t.Errorf("writeData should not return error, got: %v", err)
		}
		if !strings.Contains(buf.String(), "John,2023-01-15") {
			t.Error("Should format date according to column format")
		}
	})
}

// TestCSV_writeHeaders tests header generation
func TestCSV_writeHeaders(t *testing.T) {
	tests := []struct {
		name     string
		columns  Columns
		expected []string
	}{
		{
			name: "FlatHeaders",
			columns: Columns{
				{Name: "name", Label: "Name"},
				{Name: "age", Label: "Age"},
				{Name: "city", Label: "City"},
			},
			expected: []string{"Name,Age,City"},
		},
		{
			name: "HierarchicalHeaders",
			columns: Columns{
				{
					Label: "Group1",
					Columns: Columns{
						{Name: "col1", Label: "Col1"},
						{Name: "col2", Label: "Col2"},
					},
				},
				{Name: "col3", Label: "Col3"},
			},
			expected: []string{"Group1,,Col3", "Col1,Col2,"},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			table := &Table{Columns: tt.columns}
			buf := &bytes.Buffer{}
			csvInstance := &csv{writer: stdcsv.NewWriter(buf), table: table}

			if err := csvInstance.writeHeaders(); err != nil {
				t.Errorf("writeHeaders should not return error, got: %v", err)
			}

			csvInstance.writer.Flush()
			lines := strings.Split(strings.TrimSpace(buf.String()), "\n")

			for i, exp := range tt.expected {
				if i >= len(lines) || lines[i] != exp {
					t.Errorf("Line %d: expected '%s', got '%s'", i, exp, lines[i])
				}
			}
		})
	}
}

// TestCSV_fillHeaderLevel tests the fillHeaderLevel method
func TestCSV_fillHeaderLevel(t *testing.T) {
	tests := []struct {
		name        string
		columns     Columns
		targetLevel int
		expected    []string
		finalIndex  int
	}{
		{
			name: "FlatColumnsAtTargetLevel",
			columns: Columns{
				{Name: "col1", Label: "Col1"},
				{Name: "col2", Label: "Col2"},
				{Name: "col3", Label: "Col3"},
			},
			targetLevel: 0,
			expected:    []string{"Col1", "Col2", "Col3"},
			finalIndex:  3,
		},
		{
			name: "ParentColumnsAtTargetLevel",
			columns: Columns{
				{
					Label: "Group1",
					Columns: Columns{
						{Name: "col1", Label: "Col1"},
						{Name: "col2", Label: "Col2"},
					},
				},
				{Name: "col3", Label: "Col3"},
			},
			targetLevel: 0,
			expected:    []string{"Group1", "", "Col3"},
			finalIndex:  3,
		},
		{
			name: "LeafColumnsAtDeeperLevel",
			columns: Columns{
				{
					Label: "Group1",
					Columns: Columns{
						{Name: "col1", Label: "Col1"},
						{Name: "col2", Label: "Col2"},
					},
				},
				{Name: "col3", Label: "Col3"},
			},
			targetLevel: 1,
			expected:    []string{"Col1", "Col2", ""},
			finalIndex:  3,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			csvInstance := &csv{table: &Table{Columns: tt.columns}}
			headerRow := make([]string, len(tt.expected))

			finalIndex := csvInstance.fillHeaderLevel(headerRow, tt.targetLevel, 0, 0, tt.columns)

			if finalIndex != tt.finalIndex {
				t.Errorf("Expected final index %d, got %d", tt.finalIndex, finalIndex)
			}

			for i, exp := range tt.expected {
				if headerRow[i] != exp {
					t.Errorf("Expected '%s' at index %d, got '%s'", exp, i, headerRow[i])
				}
			}
		})
	}
}

// TestCSV_processValue tests value processing
func TestCSV_processValue(t *testing.T) {
	tests := []struct {
		name     string
		value    interface{}
		format   string
		listSep  string
		expected string
		wantErr  bool
	}{
		{"String", "test", "", "", "test", false},
		{"Integer", 42, "", "", "42", false},
		{"Boolean", true, "", "", "true", false},
		{"SliceWithSeparator", []interface{}{"a", "b", "c"}, "", "|", "a|b|c", false},
		{"DateWithFormat", time.Date(2023, 1, 15, 0, 0, 0, 0, time.UTC), "2006-01-02", "", "2023-01-15", false},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			table := &Table{ListSeparator: tt.listSep}
			csvInstance := &csv{table: table}

			result, err := csvInstance.processValue(tt.value, tt.format)

			if (err != nil) != tt.wantErr {
				t.Errorf("ProcessValue() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			if result != tt.expected {
				t.Errorf("ProcessValue() = %v, want %v", result, tt.expected)
			}
		})
	}
}
