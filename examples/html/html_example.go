// html_example.go demonstrates how to use the HTML format

package main

import (
	"log"

	"github.com/Zapharaos/go-spit"
)

func main() {
	// Sample data
	data := spit.DataSlice{
		{"name": "John Doe", "age": 30, "department": "Engineering", "site": "https://example.com"},
		{"name": "Jane Smith", "age": 28, "department": "Engineering", "site": "https://example.org"},
		{"name": "Bob Martin", "age": 41, "department": "Marketing", "site": "https://example.net"},
	}

	// Define hierarchical columns with per-column styling and a vertical merge.
	columns := spit.Columns{
		spit.NewColumn("name", "Employee Name").
			WithStyle(&spit.Style{Bold: true, TextColor: "#1155CC"}),
		spit.NewColumn("", "Details").
			WithSubColumns(spit.Columns{
				spit.NewColumn("age", "Age").
					WithStyle(&spit.Style{Alignment: spit.AlignmentRight}),
				spit.NewColumn("department", "Department").
					// Merge consecutive identical departments vertically.
					WithMerge(spit.NewMergeRules(spit.MergeConditions{spit.MergeConditionIdentical}, nil)),
			}),
		spit.NewColumn("site", "Website").
			WithFormat(spit.ExcelizeFormatHyperlink),
	}

	table := spit.NewTable(data, columns, true)

	// Document-level presentation: title, description, page font/background.
	opts := spit.HTMLOptions{
		Title:       "Employee Report",
		Description: "Generated with go-spit — HTML export example.",
		BodyStyle: &spit.Style{
			FontFamily:      "Segoe UI",
			BackgroundColor: "#FAFAFA",
			TextColor:       "#222222",
		},
		CustomCSS: "h1 { margin-bottom: 0.2em; } p { color: #666; }",
	}

	result, err := spit.ExportHTML(table, opts, spit.FileWriteParams{
		Filename:      "employee_report",
		OverwriteFile: true,
	})
	if err != nil {
		log.Fatal(err)
	}

	log.Printf("HTML written to %s", result.Filepath)
}
