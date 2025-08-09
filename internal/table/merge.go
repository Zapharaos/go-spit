package table

import (
	"fmt"
	"strings"
)

// MergeCondition defines the conditions under which cells should be merged.
// These conditions are evaluated when determining whether adjacent cells
// should be combined into a single merged cell.
type MergeCondition string

const (
	// MergeConditionIdentical merges cells when their values are identical and non-empty
	MergeConditionIdentical MergeCondition = "identical"

	// MergeConditionEmpty merges cells when both values are empty or nil
	MergeConditionEmpty MergeCondition = "empty"
)

// MergeConfig holds merge conditions for columns and rows.
// It defines when and how cells should be merged based on their content.
// Empty conditions arrays mean no merging will be applied.
type MergeConfig struct {
	Vertical   []MergeCondition `json:"vertical,omitempty"`   // Conditions for merging cells vertically (between rows)
	Horizontal []MergeCondition `json:"horizontal,omitempty"` // Conditions for merging cells horizontally (between columns)
}

// areMergeConditionsCompatible checks if two sets of merge conditions share at least one common condition.
// This is used to determine if two cells or ranges can be merged together based on their configurations.
func areMergeConditionsCompatible(conditions1, conditions2 []MergeCondition) bool {
	for _, cond1 := range conditions1 {
		for _, cond2 := range conditions2 {
			if cond1 == cond2 {
				return true // Found a matching condition
			}
		}
	}
	return false // No compatible conditions found
}

// evaluateMergeConditions determines if two values should be merged based on the specified conditions.
func evaluateMergeConditions(value1, value2 interface{}, conditions []MergeCondition) bool {
	if len(conditions) == 0 {
		return false // No conditions specified - don't merge
	}

	// Convert values to strings for consistent comparison
	val1Str := strings.TrimSpace(fmt.Sprintf("%v", value1))
	val2Str := strings.TrimSpace(fmt.Sprintf("%v", value2))

	// Determine if values are considered empty
	isEmpty1 := val1Str == "" || val1Str == "<nil>"
	isEmpty2 := val2Str == "" || val2Str == "<nil>"

	// Evaluate each condition to see if any match
	for _, condition := range conditions {
		switch condition {
		case MergeConditionIdentical:
			// Merge if values are identical and both are non-empty
			if val1Str == val2Str && !isEmpty1 && !isEmpty2 {
				return true
			}
		case MergeConditionEmpty:
			// Merge if both values are empty
			if isEmpty1 && isEmpty2 {
				return true
			}
		}
	}
	return false // No conditions matched
}
