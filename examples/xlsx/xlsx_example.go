// xlsx_example.go demonstrates advanced XLSX features including styles, borders, and merging

package main

import (
	"log"
	"time"

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
		// Simple column with style
		spit.NewColumn("name", "Employee Name").
			WithStyle(getBasicStyle()),

		// Hierarchical column with sub-columns
		spit.NewColumn("", "Personal Information").
			WithSubColumns(spit.Columns{
				spit.NewColumn("age", "Age").
					WithStyle(getBasicStyle()).
					WithBorders(spit.NewBordersBoundaries(spit.BorderStyleThin)),
				spit.NewColumn("email", "Email Address").
					WithStyle(getBasicStyle()).
					WithBorders(spit.NewBordersBoundaries(spit.BorderStyleThin)),
			}),

		spit.NewColumn("", "Work Information").
			WithSubColumns(spit.Columns{
				spit.NewColumn("department", "Department").
					WithStyle(getBasicStyle()).
					WithMerge(spit.NewMergeRules(spit.MergeConditions{spit.MergeConditionIdentical}, nil)).
					WithBorders(spit.NewBordersBoundaries(spit.BorderStyleThin).SetInner(spit.BorderStyleDouble)),
				spit.NewColumn("start_date", "Start Date").
					WithFormat("2006-01-02").
					WithStyle(getBasicStyle()).
					WithBorders(spit.NewBordersBoundaries(spit.BorderStyleThin)),
			}),
	}
}

// Sample data with various scenarios
func getSampleData() spit.DataSlice {
	return spit.DataSlice{
		{
			"name":       "John Doe",
			"age":        30,
			"email":      "john@company.com",
			"department": "Engineering",
			"start_date": time.Date(2020, 3, 15, 0, 0, 0, 0, time.UTC),
		},
		{
			"name":       "Jane Smith",
			"age":        28,
			"email":      "jane@company.com",
			"department": "Engineering",
			"start_date": time.Date(2013, 7, 29, 0, 0, 0, 0, time.UTC),
		},
		{
			"name":       "Sam Taylor",
			"age":        40,
			"email":      "sam.taylor@example.com",
			"department": "Engineering",
			"start_date": time.Date(2024, 6, 11, 0, 0, 0, 0, time.UTC),
		},
		{
			"name":       "Lisa Brown",
			"age":        27,
			"email":      "lisa.brown@example.com",
			"department": "Engineering",
			"start_date": time.Date(2019, 8, 22, 0, 0, 0, 0, time.UTC),
		},
		{
			"name":       "Bob Johnson",
			"age":        35,
			"email":      "bob@company.com",
			"department": "Marketing",
			"start_date": time.Date(2021, 1, 10, 0, 0, 0, 0, time.UTC),
		},
		{
			"name":       "Alice Brown",
			"age":        32,
			"email":      "alice@company.com",
			"department": "Engineering",
			"start_date": time.Date(2018, 11, 5, 0, 0, 0, 0, time.UTC),
		},
		{
			"name":       "N/A",
			"age":        0,
			"email":      "N/A",
			"department": "N/A",
			"start_date": "N/A",
		},
	}
}

// Create row options using the new fluent API
func getRowOptions() spit.RowOptionsMap {
	return spit.RowOptionsMap{
		1: *spit.NewRowOptions(1).
			WithStyle(&spit.Style{
				BackgroundColor: "#FFE6E6",
				Alignment:       spit.AlignmentCenterMiddle,
			}).
			WithBorder(spit.NewBordersBoundaries(spit.BorderStyleThick)),

		6: *spit.NewRowOptions(6).
			WithStyle(&spit.Style{
				BackgroundColor: "#FFE6E6",
				Alignment:       spit.AlignmentCenterMiddle,
			}).
			WithMerge(spit.NewMergeRules(nil, spit.MergeConditions{spit.MergeConditionIdentical})).
			WithBorder(spit.NewBordersBoundaries(spit.BorderStyleThin)),
	}
}

// Create cell options using the new fluent API
func getCellOptions() spit.CellOptionsMap {
	return spit.CellOptionsMap{
		4: { // Column index 4 (department column)
			5: *spit.NewCellOptions(5, 4).
				WithStyle(&spit.Style{
					BackgroundColor: "#FFFF99",
					Alignment:       spit.AlignmentCenterMiddle,
				}).
				WithBorder(spit.NewBordersBoundaries(spit.BorderStyleThick)).
				WithMergeable(false),
		},
		2: { // Column index 2 (department column)
			6: *spit.NewCellOptions(6, 2).
				WithStyle(&spit.Style{
					BackgroundColor: "#FFFF99",
					Alignment:       spit.AlignmentCenterMiddle,
				}).
				WithBorder(spit.NewBordersBoundaries(spit.BorderStyleThick)).
				WithMergeable(false),
		},
	}
}

func main() {
	// Get sample data and columns
	data := getSampleData()
	columns := getMainColumns()

	// Create table using the new constructor and fluent API
	table := spit.NewTable(data, columns, true).
		WithRowOptions(getRowOptions()).
		WithCellOptions(getCellOptions())

	// Create spreadsheet
	spreadsheet := spit.NewSpreadsheetExcelize("Employee Report", table)

	// Export with advanced file options
	params := spit.FileWriteParams{
		Filename:      "advanced_employee_report",
		UseTempFile:   false,
		OverwriteFile: true,
	}

	result, err := spit.ExportXLSX(spreadsheet, params)
	if err != nil {
		log.Fatalf("Error writing XLSX file: %v", err)
	}

	log.Printf("Successfully created XLSX file: %s", result.Filepath)

	// Uncomment to remove the file after creation
	// defer func() {
	//     if closeErr := result.RemoveFile(); closeErr != nil {
	//         log.Printf("Failed to remove XLSX file: %v", closeErr)
	//     }
	// }()
}
