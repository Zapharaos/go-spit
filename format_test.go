package spit

import (
	"testing"
)

func TestFormat_String(t *testing.T) {
	tests := []struct {
		format   Format
		expected string
	}{
		{FormatCSV, "csv"},
		{FormatUnknown, "Format(0)"},
		{Format(99), "Format(99)"},
	}

	for _, tt := range tests {
		result := tt.format.String()
		if result != tt.expected {
			t.Errorf("Format(%d).String() = %q, want %q", tt.format, result, tt.expected)
		}
	}
}
