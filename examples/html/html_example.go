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

	log.Printf("HTML table written to %s", result.Filepath)

	// A full document composes headings, paragraphs, lists, sections and tables.
	// The built-in theme gives it a polished look, and TableOfContents adds anchors.
	doc := spit.NewHTMLDocument(spit.HTMLOptions{
		Title:           "Employee Report",
		Description:     "Generated with go-spit — full document example.",
		Theme:           spit.HTMLThemeDefault,
		TableOfContents: true,
	}).
		Heading(2, "Summary").
		Paragraph("This report lists the current engineering and marketing staff.").
		DefinitionList(
			spit.Def("Prepared by", "People Ops"),
			spit.Def("Period", "July 2026"),
		).
		Add(spit.UnorderedList("Two departments").
			Add(spit.Item("Engineering", spit.Item("Backend"), spit.Item("Frontend"))).
			Add(spit.Item("Marketing"))).
		Section(2, "Directory",
			spit.Paragraph("Full employee directory:"),
			spit.TableBlock(table).WithCaption("Current staff"),
		).
		HorizontalRule().
		Blockquote("Generated automatically — do not edit by hand.")

	docResult, err := spit.ExportHTMLDocument(doc, spit.FileWriteParams{
		Filename:      "employee_document",
		OverwriteFile: true,
	})
	if err != nil {
		log.Fatal(err)
	}

	log.Printf("HTML document written to %s", docResult.Filepath)
}
