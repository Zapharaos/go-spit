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
		return v.Format(format), nil
	case string:
		date, err := ParseDate(v)
		if err == nil {
			return date.Format(format), nil
		}
		return nil, err
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
