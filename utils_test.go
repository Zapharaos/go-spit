package spit

import (
	"testing"
	"time"
)

func TestConvertSliceToString(t *testing.T) {
	tests := []struct {
		name     string
		slice    []interface{}
		format   string
		sep      string
		expected string
		wantErr  bool
	}{
		{
			name:     "With date formatting - mixed types (no error since strings skip formatting)",
			slice:    []interface{}{time.Date(2024, 6, 1, 12, 34, 56, 789000000, time.UTC), "foo", 42},
			format:   "2006-01-02",
			sep:      ",",
			expected: "2024-06-01,foo,42",
			wantErr:  false,
		},
		{
			name:     "With date formatting - only time values",
			slice:    []interface{}{time.Date(2024, 6, 1, 12, 34, 56, 789000000, time.UTC), time.Date(2024, 12, 25, 10, 30, 45, 123000000, time.UTC)},
			format:   "2006-01-02",
			sep:      ",",
			expected: "2024-06-01,2024-12-25",
			wantErr:  false,
		},
		{
			name:     "Without formatting",
			slice:    []interface{}{"2024-06-01T12:34:56.789", "foo", 42},
			format:   "",
			sep:      "|",
			expected: "2024-06-01T12:34:56.789|foo|42",
			wantErr:  false,
		},
		{
			name:     "Empty slice",
			slice:    []interface{}{},
			format:   "",
			sep:      ",",
			expected: "",
			wantErr:  false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result, err := ConvertSliceToString(tt.slice, tt.format, tt.sep)
			if tt.wantErr {
				if err == nil {
					t.Fatalf("expected error, got nil")
				}
			} else {
				if err != nil {
					t.Fatalf("unexpected error: %v", err)
				}
				if result != tt.expected {
					t.Errorf("got %q, want %q", result, tt.expected)
				}
			}
		})
	}
}

func TestFormatValue(t *testing.T) {
	tests := []struct {
		name     string
		value    interface{}
		format   string
		expected interface{}
		wantErr  bool
	}{
		{
			name:     "Time value",
			value:    time.Date(2024, 6, 1, 12, 34, 56, 789000000, time.UTC),
			format:   "2006-01-02",
			expected: "2024-06-01",
			wantErr:  false,
		},
		{
			name:     "String date value - should not be formatted",
			value:    "2024-06-01T12:34:56.789",
			format:   "2006-01-02",
			expected: "2024-06-01T12:34:56.789", // Strings skip formatting
			wantErr:  false,
		},
		{
			name:     "String non-date value - should not be formatted",
			value:    "notadate",
			format:   "2006-01-02",
			expected: "notadate", // Strings skip formatting
			wantErr:  false,
		},
		{
			name:     "Time pointer value",
			value:    func() *time.Time { t := time.Date(2024, 6, 1, 12, 34, 56, 789000000, time.UTC); return &t }(),
			format:   "2006-01-02",
			expected: "2024-06-01",
			wantErr:  false,
		},
		{
			name:     "Nil time pointer",
			value:    (*time.Time)(nil),
			format:   "2006-01-02",
			expected: "",
			wantErr:  false,
		},
		{
			name:     "Integer value",
			value:    123,
			format:   "unused",
			expected: 123,
			wantErr:  false,
		},
		{
			name:     "Float value",
			value:    123.45,
			format:   "unused",
			expected: 123.45,
			wantErr:  false,
		},
		{
			name:     "Boolean value",
			value:    true,
			format:   "unused",
			expected: true,
			wantErr:  false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			val, err := FormatValue(tt.value, tt.format)
			if tt.wantErr {
				if err == nil {
					t.Errorf("expected error, got nil")
				}
			} else {
				if err != nil {
					t.Fatalf("unexpected error: %v", err)
				}
				if val != tt.expected {
					t.Errorf("got %v, want %v", val, tt.expected)
				}
			}
		})
	}
}

func TestParseDate(t *testing.T) {
	tests := []struct {
		name     string
		input    string
		wantYear int
		wantErr  bool
	}{
		{
			name:     "Valid date format 1",
			input:    "2024-06-01T12:34:56.789",
			wantYear: 2024,
			wantErr:  false,
		},
		{
			name:     "Valid date format 2",
			input:    "2024-06-01T12:4:56.789",
			wantYear: 2024,
			wantErr:  false,
		},
		{
			name:     "Invalid date",
			input:    "invalid",
			wantYear: 0,
			wantErr:  true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			date, err := ParseDate(tt.input)
			if tt.wantErr {
				if err == nil {
					t.Errorf("ParseDate(%q) expected error, got nil", tt.input)
				}
			} else {
				if err != nil {
					t.Errorf("ParseDate(%q) unexpected error: %v", tt.input, err)
				}
				if date.Year() != tt.wantYear {
					t.Errorf("ParseDate(%q) got year %d, want %d", tt.input, date.Year(), tt.wantYear)
				}
			}
		})
	}
}
