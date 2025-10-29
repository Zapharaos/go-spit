// utils.go - Utility functions.
//
// This file provides helper functions, designed to be generic for different export formats.

package spit

import (
	"fmt"
	"strings"
	"time"
)

// ConvertSliceToString takes a slice []interface{}, formats each element, and returns the elements separated by a custom separator.
// Used for rendering array/slice values in exported files.
func ConvertSliceToString(slice []interface{}, format string, separator string) (string, error) {
	var strValues []string
	for _, elem := range slice {
		if format != "" {
			var err error
			elem, err = FormatValue(elem, format)
			if err != nil {
				return "", err
			}
		}
		strValues = append(strValues, fmt.Sprintf("%v", elem))
	}
	return strings.Join(strValues, separator), nil
}

// FormatValue applies the specified format to a given value.
// Supports time.Time and string values that can be parsed as dates.
func FormatValue(value interface{}, format string) (interface{}, error) {
	switch v := value.(type) {
	case time.Time:
		if format != "" {
			return v.Format(format), nil
		}
		return v, nil
	case *time.Time:
		if v != nil {
			if format != "" {
				return v.Format(format), nil
			}
			return *v, nil
		}
		return "", nil
	case string:
		// Skip formatting for string values, even if format is specified
		// This prevents format conflicts (e.g., "Total" being formatted as date)
		return v, nil
	case int, int8, int16, int32, int64, uint, uint8, uint16, uint32, uint64, float32, float64, bool:
		return v, nil
	}
	return value, nil
}

// ParseDate parses a date string using several supported formats.
// Returns a time.Time if parsing succeeds, or an error otherwise.
func ParseDate(dateStr string) (time.Time, error) {
	dateFormats := []string{
		"2006-01-02T15:04:05.999",
		"2006-01-02T15:4:05.999",
		"2006-01-02T15:04:5.999",
		"2006-01-02T15:4:5.999",
	}

	for _, format := range dateFormats {
		if date, err := time.Parse(format, dateStr); err == nil {
			return date, nil
		}
	}

	return time.Time{}, fmt.Errorf("failed to parse date string: %s", dateStr)
}

// parseAsInt attempts to parse a string as an integer.
// Returns the parsed int64 value or an error if parsing fails.
func parseAsInt(s string) (int64, error) {
	s = strings.TrimSpace(s)
	if s == "" {
		return 0, fmt.Errorf("empty string")
	}

	var result int64
	_, err := fmt.Sscanf(s, "%d", &result)
	return result, err
}

// parseAsFloat attempts to parse a string as a floating-point number.
// Returns the parsed float64 value or an error if parsing fails.
func parseAsFloat(s string) (float64, error) {
	s = strings.TrimSpace(s)
	if s == "" {
		return 0, fmt.Errorf("empty string")
	}

	var result float64
	_, err := fmt.Sscanf(s, "%f", &result)
	return result, err
}

// parseAsBool attempts to parse a string as a boolean.
// Recognizes common boolean string representations (true/false, yes/no, 1/0).
func parseAsBool(s string) (bool, error) {
	s = strings.TrimSpace(strings.ToLower(s))
	switch s {
	case "true", "yes", "1", "t", "y":
		return true, nil
	case "false", "no", "0", "f", "n":
		return false, nil
	default:
		return false, fmt.Errorf("cannot parse '%s' as boolean", s)
	}
}
