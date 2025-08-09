package utils

import (
	"fmt"
	"strings"
	"time"

	"github.com/Zapharaos/go-spit/internal/logger"
)

// ConvertSliceToString Take a slice []interface{}, format and returns the elements separated by custom separator.
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

// FormatValue Apply the specified format to a given value.
func FormatValue(value interface{}, format string) (interface{}, error) {
	switch v := value.(type) {
	case time.Time:
		return v.Format(format), nil
	case string:
		date, err := parseDate(v)
		if err == nil {
			return date.Format(format), nil
		}
		logger.L().Error("Failed to parse date string:", logger.Any("value", v), logger.Error(err))
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
