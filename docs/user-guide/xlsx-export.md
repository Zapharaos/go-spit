# XLSX Export

go-spit exports rich Excel spreadsheets through a `Spreadsheet` implementation. The library ships
with an [Excelize](https://github.com/xuri/excelize)-backed implementation,
`SpreadsheetExcelize`.

## The export functions

```go
// Export a single sheet.
func ExportXLSX(s Spreadsheet, params FileWriteParams) (*FileWriteResult, error)

// Export one or more sheets into a single workbook.
func ExportXLSXSheets(sheets []Spreadsheet, params FileWriteParams) (*FileWriteResult, error)
```

The `.xlsx` extension is added automatically when `params.Extension` is empty. See
[File Options](file-options.md) for the available `params`.

## Basic example

```go
package main

import (
	"log"

	"github.com/Zapharaos/go-spit"
)

func main() {
	data := spit.DataSlice{
		{"name": "John Doe", "age": 30, "department": "Engineering"},
		{"name": "Jane Smith", "age": 28, "department": "Marketing"},
	}

	columns := spit.Columns{
		spit.NewColumn("name", "Employee Name"),
		spit.NewColumn("", "Details").
			WithSubColumns(spit.Columns{
				spit.NewColumn("age", "Age"),
				spit.NewColumn("department", "Department"),
			}),
	}

	table := spit.NewTable(data, columns, true)
	spreadsheet := spit.NewSpreadsheetExcelize("Employee Report", table)

	result, err := spit.ExportXLSX(spreadsheet, spit.FileWriteParams{
		Filename: "report",
	})
	if err != nil {
		log.Fatal(err)
	}
	defer result.RemoveFile()

	log.Printf("created %s", result.Filepath)
}
```

`NewSpreadsheetExcelize(sheetName, table)` creates a spreadsheet bound to a sheet name and a
table. When no underlying file exists yet, `ExportXLSX` creates a fresh workbook for you.

## Multiple sheets

To write several sheets into one workbook, build a `Spreadsheet` per sheet and pass them to
`ExportXLSXSheets`. The first sheet's workbook is shared with the others automatically:

```go
sheet1 := spit.NewSpreadsheetExcelize("Engineering", engineeringTable)
sheet2 := spit.NewSpreadsheetExcelize("Marketing", marketingTable)

result, err := spit.ExportXLSXSheets(
	[]spit.Spreadsheet{sheet1, sheet2},
	spit.FileWriteParams{Filename: "departments"},
)
if err != nil {
	log.Fatal(err)
}
defer result.RemoveFile()
```

## Using an existing Excelize file

If you already have an `*excelize.File` (for instance to add go-spit sheets to a pre-built
workbook), attach it with `WithFile`:

```go
f := excelize.NewFile()
spreadsheet := spit.NewSpreadsheetExcelize("Report", table).WithFile(f)
```

## Cell content formats

The `Format` field on a column controls how XLSX cell content is written. In addition to date
layouts, go-spit defines Excelize-specific constants:

| Constant                  | Value         | Behavior                                                                    |
|---------------------------|---------------|------------------------------------------------------------------------------|
| `ExcelizeFormatDefault`   | `"default"`   | Passes the raw value to Excelize, preserving the native Go type.            |
| `ExcelizeFormatFormula`   | `"formula"`   | Writes the value as an Excel formula, e.g. `"=SUM(A1:A10)"`.                 |
| `ExcelizeFormatHyperlink` | `"hyperlink"` | Writes a clickable external hyperlink; the value must be a URL string.       |

```go
columns := spit.Columns{
	spit.NewColumn("total", "Total").WithFormat(spit.ExcelizeFormatFormula),
	spit.NewColumn("homepage", "Website").WithFormat(spit.ExcelizeFormatHyperlink),
}
```

## Advanced features

XLSX export supports the full styling model:

- **Styles** — font family, size, colors, bold, italic and alignment.
- **Borders** — thin, medium, thick, double, dashed and dotted borders, with inner-border support.
- **Cell merging** — vertical and horizontal merging based on identical or empty values.
- **Row options** — apply styling, borders and merging to entire rows.
- **Cell options** — fine-grained styling, borders and merging for individual cells.
- **Column formatting** — dates, formulas, hyperlinks and custom value formats.

These are covered in detail in [Styling, Borders & Merging](styling.md).

## The Spreadsheet interface

`Spreadsheet` abstracts spreadsheet operations so additional backends can be implemented. The
`SpreadsheetExcelize` type is the default implementation. Implementing the interface yourself lets
you target other spreadsheet libraries while reusing the rest of go-spit. See the
[API Reference](../reference/api.md) for the full method set.
