// xlsx_example.go demonstrates advanced XLSX features including styles, borders, and merging

package main

import (
	"log"

	"github.com/Zapharaos/go-spit"
)

// Simple style functions
func getBasicStyle() *spit.Style {
	return &spit.Style{
		Alignment: spit.AlignmentCenterMiddle,
	}
}

// Column definitions
func getMainColumns() spit.Columns {
	return spit.Columns{
		// Column without subcolumns, no borders
		{
			Name:  "name",
			Label: "Name",
			Style: getBasicStyle(),
		},
		// Column with subcolumns, subcolumns with dedicated borders
		{
			Label: "Personal Info",
			Columns: spit.Columns{
				{
					Name:    "age",
					Label:   "Age",
					Style:   getBasicStyle(),
					Borders: spit.NewBorderOptions(spit.BorderStyleThin),
				},
				{
					Name:    "email",
					Label:   "Email",
					Style:   getBasicStyle(),
					Borders: spit.NewBorderOptions(spit.BorderStyleThin),
				},
			},
		},
		// Column with subcolumns and inner borders
		{
			Label:   "Work Info",
			Borders: spit.NewBorderOptions(spit.BorderStyleThin),
			Columns: spit.Columns{
				{
					Name:    "department",
					Label:   "Department",
					Style:   getBasicStyle(),
					Borders: spit.NewBorderOptions(spit.BorderStyleThin).SetInner(spit.BorderStyleDouble),
					Merge: &spit.MergeRules{
						Vertical: spit.MergeConditions{spit.MergeConditionIdentical},
					}, // Show vertical merge
				},
				{
					Name:    "status",
					Label:   "Status",
					Style:   getBasicStyle(),
					Borders: spit.NewBorderOptions(spit.BorderStyleThin).SetInner(spit.BorderStyleDashed),
				},
			},
		},
	}
}

// Simple sample data
func getSampleData() spit.DataSlice {
	return spit.DataSlice{
		{
			"name":       "John Doe",
			"age":        30,
			"email":      "john@example.com",
			"department": "Engineering",
			"status":     "Active",
		},
		{
			"name":       "Jane Smith",
			"age":        28,
			"email":      "jane@example.com",
			"department": "Engineering",
			"status":     "Active",
		},
		{
			"name":       "Sam Taylor",
			"age":        40,
			"email":      "sam.taylor@example.com",
			"department": "Engineering",
			"status":     "Disabled",
		},
		{
			"name":       "Lisa Brown",
			"age":        27,
			"email":      "lisa.brown@example.com",
			"department": "Engineering",
			"status":     "Disabled",
		},
		{
			"name":       "Bob Johnson",
			"age":        35,
			"email":      "bob@example.com",
			"department": "Marketing",
			"status":     "Active",
		},
		{
			"name":       "Alice Wilson",
			"age":        32,
			"email":      "alice@example.com",
			"department": "Marketing",
			"status":     "Active",
		},
		{
			"name":       "N/A",
			"age":        0,
			"email":      "N/A",
			"department": "N/A",
			"status":     "N/A",
		},
	}
}

// Minimal row options - just background color
func getRowOptions() spit.RowOptionsMap {
	return spit.RowOptionsMap{
		1: spit.RowOptions{
			RowIndex: 1,
			Style: &spit.Style{
				BackgroundColor: "#FFE6E6",
				Alignment:       spit.AlignmentCenterMiddle,
			},
			Border:    spit.NewBorderOptions(spit.BorderStyleThick),
			Mergeable: false, // Cancel any column merge behavior for this row
		},
		6: spit.RowOptions{
			RowIndex: 6,
			Style: &spit.Style{
				BackgroundColor: "#FFE6E6",
				Alignment:       spit.AlignmentCenterMiddle,
			},
			Merge: &spit.MergeRules{
				Horizontal: spit.MergeConditions{spit.MergeConditionIdentical},
			}, // Show horizontal merge
			Border: spit.NewBorderOptions(spit.BorderStyleThin),
		},
	}
}

// Minimal cell options - just background color
func getCellOptions() spit.CellOptionsMap {
	return spit.CellOptionsMap{
		4: { // Column index 4 (department column)
			5: spit.CellOptions{ // Row index 5 (sixth data row)
				RowIndex: 5,
				ColIndex: 4,
				Style: &spit.Style{
					BackgroundColor: "#FFFF99",
					Alignment:       spit.AlignmentCenterMiddle,
				},
				Mergeable: false, // Refuses to merge this cell
				Border:    spit.NewBorderOptions(spit.BorderStyleThick),
			},
		},
		2: { // Column index 2 (department column)
			6: spit.CellOptions{ // Row index 6 (seventh data row)
				RowIndex: 6,
				ColIndex: 2,
				Style: &spit.Style{
					BackgroundColor: "#FFFF99",
					Alignment:       spit.AlignmentCenterMiddle,
				},
				Mergeable: false, // Refuses to merge this cell
				Border:    spit.NewBorderOptions(spit.BorderStyleThick),
			},
		},
	}
}

func main() {
	// Create a simple table with focused features
	table := &spit.Table{
		Data:           getSampleData(),
		Columns:        getMainColumns(),
		RowOptionsMap:  getRowOptions(),
		CellOptionsMap: getCellOptions(),
		WriteHeader:    true,
		Limit:          0,
	}

	// File parameters
	params := spit.FileWriteParams{
		Filename:    "simplified_xlsx_example",
		UseTempFile: true,
	}

	// Create and export
	spreadsheet := spit.NewSpreadsheetExcelize("Simple XLSX Demo", table)
	result, err := spit.ExportXLSX(spreadsheet, params)
	if err != nil {
		log.Fatalf("Error writing XLSX file: %v", err)
	}

	defer func() {
		if closeErr := result.RemoveFile(); closeErr != nil {
			log.Printf("Failed to remove XLSX file: %v", closeErr)
		}
	}()
}
