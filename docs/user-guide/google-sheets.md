# Google Sheets

go-spit can export tables straight to **Google Sheets**, reusing the same table model, styling and
cell merging as the file backends.

Because the Google API client pulls in a large dependency tree, this lives in a **separate,
optional module** so the core `go-spit` stays dependency-light. Only projects that need Sheets pull
it in:

```sh
go get github.com/Zapharaos/go-spit/gsheets
```

## Authentication stays with you

The `gsheets` package **never handles credentials**. You build an authenticated `*sheets.Service`
(service account, OAuth, or Application Default Credentials — with the
`https://www.googleapis.com/auth/spreadsheets` scope) and pass it in. Nothing is bundled, and there
is no external service or Docker involved.

## Example

```go
ctx := context.Background()

// You own authentication; the library only receives the ready-to-use service.
svc, err := sheets.NewService(ctx, option.WithCredentialsFile("service-account.json"))
if err != nil {
    log.Fatal(err)
}

table := spit.NewTable(data, columns, true)

// Empty spreadsheetID creates a new spreadsheet; pass an ID to update an existing one.
result, err := gsheets.ExportGoogleSheets(ctx, svc, "", []gsheets.Sheet{
    {Name: "Employees", Table: table},
}, gsheets.Options{Title: "Staff"})
if err != nil {
    log.Fatal(err)
}
log.Printf("written to %s", result.URL)
```

Provide multiple `Sheet` entries to write several tabs in one call.

## How it maps

The backend implements go-spit's `TableOperations` interface, so merging and styling come from the
shared pipelines. Everything is applied in a single `batchUpdate` per export.

| go-spit | Google Sheets |
|---------|---------------|
| cell value | `ExtendedValue` (numbers/booleans stay native) |
| cell merging | `MergeCellsRequest` |
| `Style` | `CellFormat` (text format, background, alignment) |
| `Borders` | `CellFormat.Borders` |
| `NumFmt` | `CellFormat.NumberFormat` |
| `ExcelizeFormatHyperlink` | `=HYPERLINK(...)` |
| `Image` (URL) | `=IMAGE(...)` |

## Caveats

- It is a **network push**, not a file: mind `ctx` timeouts and the Sheets API **rate limits**
  (~60 writes/min/user). Exports are batched to keep the request count low.
- Only **URL images** are supported (`=IMAGE`); embedded `Image.Bytes` fall back to alt text.

Full details and the module README are in
[`gsheets/`](https://github.com/Zapharaos/go-spit/tree/main/gsheets).
