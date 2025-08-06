package go_spit

import (
	"fmt"
	"sort"
	"strings"
	"time"
)

type Data map[string]interface{}

type DataSlice []Data

// Lookup looks up a key in a nested map structure.
func (d Data) Lookup(ks ...string) (rval interface{}, err error) {
	var ok bool
	if len(ks) == 0 {
		return nil, fmt.Errorf("nestedMapLookup needs at least one key")
	}
	if rval, ok = d[strings.TrimSpace(ks[0])]; !ok {
		return nil, fmt.Errorf("key not found; remaining keys: %v", ks)
	} else if len(ks) == 1 {
		return rval, nil
	} else if d, ok = rval.(map[string]interface{}); !ok {
		return nil, fmt.Errorf("malformed structure at %#v", rval)
	} else {
		return d.Lookup(ks[1:]...)
	}
}

// SortByTime sorts data by key.
func (d DataSlice) SortByTime(key string) DataSlice {
	if len(d) == 0 {
		return d
	}

	sorted := make(DataSlice, len(d))
	copy(sorted, d)

	sort.Slice(sorted, func(i, j int) bool {
		timeI, okI := sorted[i][key].(time.Time)
		timeJ, okJ := sorted[j][key].(time.Time)
		if okI && okJ {
			return timeI.Before(timeJ)
		}
		return false
	})

	return sorted
}

type Format uint8

const (
	FormatUnknown Format = iota
	FormatCSV
	FormatXSLX
)

var formats = map[Format]string{
	FormatCSV:  "csv",
	FormatXSLX: "xlsx",
}

// String returns the string representation of the Format
func (f Format) String() string {
	if str, ok := formats[f]; ok {
		return str
	}
	return fmt.Sprintf("Format(%d)", f)
}
