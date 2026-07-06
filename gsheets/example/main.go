// Command example demonstrates exporting a go-spit table to Google Sheets.
//
// Usage:
//
//	# Using a service account key file:
//	go run ./example -credentials=service-account.json
//
//	# Update an existing spreadsheet the account can edit (recommended, see note below):
//	go run ./example -credentials=service-account.json -spreadsheet=<SPREADSHEET_ID>
//
//	# Or rely on Application Default Credentials (GOOGLE_APPLICATION_CREDENTIALS / gcloud):
//	go run ./example
//
// Note on service accounts: a spreadsheet CREATED by a service account is owned by that
// service account's Drive and won't appear in your own Google Drive UI. For easy testing,
// create a blank spreadsheet in your Drive, share it (Editor) with the service account
// email, and pass its ID via -spreadsheet. With OAuth user credentials this isn't an issue.
package main

import (
	"context"
	"flag"
	"log"

	spit "github.com/Zapharaos/go-spit"
	"github.com/Zapharaos/go-spit/gsheets"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

func main() {
	credentials := flag.String("credentials", "", "path to a service account JSON key (optional; falls back to Application Default Credentials)")
	spreadsheetID := flag.String("spreadsheet", "", "existing spreadsheet ID to update (optional; a new spreadsheet is created when empty)")
	flag.Parse()

	ctx := context.Background()

	// You build the authenticated service — gsheets never touches your credentials.
	opts := []option.ClientOption{option.WithScopes(sheets.SpreadsheetsScope)}
	if *credentials != "" {
		opts = append(opts, option.WithCredentialsFile(*credentials))
	}
	svc, err := sheets.NewService(ctx, opts...)
	if err != nil {
		log.Fatalf("failed to create Sheets service: %v", err)
	}

	// A sample table showcasing headers, per-column styling, a vertical merge,
	// native numbers and a hyperlink column.
	data := spit.DataSlice{
		{"name": "John Doe", "dept": "Engineering", "salary": 75000, "profile": "https://example.com/john"},
		{"name": "Jane Smith", "dept": "Engineering", "salary": 82000, "profile": "https://example.com/jane"},
		{"name": "Bob Martin", "dept": "Marketing", "salary": 68000, "profile": "https://example.com/bob"},
	}
	columns := spit.Columns{
		spit.NewColumn("name", "Employee").
			WithStyle(&spit.Style{Bold: true, TextColor: "#1155CC"}),
		spit.NewColumn("dept", "Department").
			// Merge consecutive identical departments vertically.
			WithMerge(spit.NewMergeRules(spit.MergeConditions{spit.MergeConditionIdentical}, nil)),
		spit.NewColumn("salary", "Salary"),
		spit.NewColumn("profile", "Profile").
			WithFormat(spit.ExcelizeFormatHyperlink),
	}
	table := spit.NewTable(data, columns, true)

	result, err := gsheets.ExportGoogleSheets(ctx, svc, *spreadsheetID, []gsheets.Sheet{
		{Name: "Employees", Table: table},
	}, gsheets.Options{Title: "go-spit demo"})
	if err != nil {
		log.Fatalf("export failed: %v", err)
	}

	log.Printf("done: %s", result.URL)
}
