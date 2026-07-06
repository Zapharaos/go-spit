package gsheets_test

import (
	"context"
	"log"

	spit "github.com/Zapharaos/go-spit"
	"github.com/Zapharaos/go-spit/gsheets"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

// Example shows the full flow: the caller builds an authenticated *sheets.Service and
// passes it to ExportGoogleSheets, which never handles credentials itself.
func Example() {
	ctx := context.Background()

	// The caller owns authentication (service account, OAuth, or ADC).
	svc, err := sheets.NewService(ctx, option.WithCredentialsFile("service-account.json"))
	if err != nil {
		log.Fatal(err)
	}

	table := spit.NewTable(
		spit.DataSlice{
			{"name": "John Doe", "dept": "Engineering"},
			{"name": "Jane Smith", "dept": "Engineering"},
		},
		spit.Columns{
			spit.NewColumn("name", "Employee").WithStyle(&spit.Style{Bold: true}),
			spit.NewColumn("dept", "Department").
				WithMerge(spit.NewMergeRules(spit.MergeConditions{spit.MergeConditionIdentical}, nil)),
		},
		true,
	)

	// Empty spreadsheetID creates a new spreadsheet.
	result, err := gsheets.ExportGoogleSheets(ctx, svc, "", []gsheets.Sheet{
		{Name: "Employees", Table: table},
	}, gsheets.Options{Title: "Staff"})
	if err != nil {
		log.Fatal(err)
	}

	log.Printf("written to %s", result.URL)
}
