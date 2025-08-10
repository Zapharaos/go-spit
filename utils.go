package go_spit

import (
	"fmt"
	"strings"
	"time"
)

// convertSliceToString Take a slice []interface{}, format and returns the elements separated by custom separator.
func convertSliceToString(slice []interface{}, format string, separator string) (string, error) {
	var strValues []string
	for _, elem := range slice {
		if format != "" {
			var err error
			elem, err = formatValue(elem, format)
			if err != nil {
				return "", err
			}
		}
		strValues = append(strValues, fmt.Sprintf("%v", elem))
	}
	return strings.Join(strValues, separator), nil
}

// formatValue Apply the specified format to a given value.
func formatValue(value interface{}, format string) (interface{}, error) {
	switch v := value.(type) {
	case time.Time:
		return v.Format(format), nil
	case string:
		date, err := parseDate(v)
		if err == nil {
			return date.Format(format), nil
		}
		return nil, err
	}
	return value, nil
}

// parseDate parses a date string
func parseDate(dateStr string) (time.Time, error) {
	formats := []string{
		"2006-01-02T15:04:05.999",
		"2006-01-02T15:4:05.999",
		"2006-01-02T15:04:5.999",
		"2006-01-02T15:4:5.999",
	}

	for _, format := range formats {
		if date, err := time.Parse(format, dateStr); err == nil {
			return date, nil
		}
	}

	return time.Time{}, fmt.Errorf("failed to parse date string: %s", dateStr)
}
