# go-spit / gsheets

Export [go-spit](https://github.com/Zapharaos/go-spit) tables to **Google Sheets**.

This is an **optional, separately-versioned module**: the core `go-spit` library stays
dependency-light, and only projects that import this package pull in the Google API client.

```sh
go get github.com/Zapharaos/go-spit/gsheets
```

## Authentication is the caller's responsibility

This package **never handles credentials**. You build an authenticated `*sheets.Service`
however suits your app (service account, OAuth, or Application Default Credentials) and pass it
in. The service needs the `https://www.googleapis.com/auth/spreadsheets` scope.

## Usage

```go
package main

import (
	"context"
	"log"

	spit "github.com/Zapharaos/go-spit"
	"github.com/Zapharaos/go-spit/gsheets"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

func main() {
	ctx := context.Background()

	// You provide the authenticated service — the library never sees your credentials.
	svc, err := sheets.NewService(ctx, option.WithCredentialsFile("service-account.json"))
	if err != nil {
		log.Fatal(err)
	}

	data := spit.DataSlice{
		{"name": "John Doe", "dept": "Engineering", "salary": 75000},
		{"name": "Jane Smith", "dept": "Engineering", "salary": 82000},
	}
	columns := spit.Columns{
		spit.NewColumn("name", "Employee").WithStyle(&spit.Style{Bold: true}),
		spit.NewColumn("dept", "Department").
			WithMerge(spit.NewMergeRules(spit.MergeConditions{spit.MergeConditionIdentical}, nil)),
		spit.NewColumn("salary", "Salary"),
	}
	table := spit.NewTable(data, columns, true)

	// Empty spreadsheetID => create a new spreadsheet titled "Staff".
	result, err := gsheets.ExportGoogleSheets(ctx, svc, "", []gsheets.Sheet{
		{Name: "Employees", Table: table},
	}, gsheets.Options{Title: "Staff"})
	if err != nil {
		log.Fatal(err)
	}

	log.Printf("written to %s", result.URL)
}
```

Pass an existing `spreadsheetID` to update a specific spreadsheet (missing tabs are created).
Provide multiple `Sheet` entries to write several tabs in one call.

## Runnable example

A ready-to-run example lives in [`example/`](example):

```sh
# with a service account key file
go run ./example -credentials=service-account.json

# update an existing spreadsheet the account can edit (recommended for testing)
go run ./example -credentials=service-account.json -spreadsheet=<SPREADSHEET_ID>

# or rely on Application Default Credentials
go run ./example
```

> **Testing tip:** a spreadsheet *created* by a service account is owned by that account's Drive and
> won't show up in your own Drive UI. For easy testing, create a blank spreadsheet in your Drive,
> share it (Editor) with the service account email, and pass its ID via `-spreadsheet`.

## How it works

The export implements go-spit's `TableOperations` interface, so it reuses the exact same
merging and styling pipelines as the CSV/XLSX/HTML backends. The in-memory grid is translated
into a single Google Sheets `batchUpdate` per export.

| go-spit | Google Sheets API |
|---------|-------------------|
| cell value | `UpdateCells` / `ExtendedValue` (numbers & booleans stay native) |
| cell merging | `MergeCellsRequest` + `GridRange` |
| `Style` (bold, colors, alignment, font) | `CellFormat` (`TextFormat`, `BackgroundColor`, alignment) |
| `Borders` | `CellFormat.Borders` |
| `NumFmt` | `CellFormat.NumberFormat` |
| `ExcelizeFormatFormula` | formula cell |
| `ExcelizeFormatHyperlink` | `=HYPERLINK(...)` |
| `Image` (URL) | `=IMAGE(...)` |

## Caveats

- **Network, not a file.** Respect `ctx` timeouts and the Sheets API **quotas / rate limits**
  (roughly 60 write requests per minute per user). Everything is batched into one call per export
  to minimize requests.
- **Images:** only URL images are supported (`=IMAGE`). Embedded `Image.Bytes` cannot live in a
  Sheets cell and fall back to the alt text.
- **Best effort:** exotic formats and very fine border styles map approximately, as with the other
  backends.

## Local development

This module uses a `replace github.com/Zapharaos/go-spit => ../` directive so it builds against
the repository root during development. Consumers ignore that directive and use the published
`go-spit` version.
