# CSV Export

go-spit exports tabular data to CSV with `ExportCSV`:

```go
func ExportCSV(separator string, t *Table, params FileWriteParams) (*FileWriteResult, error)
```

- **`separator`** — the field delimiter. Pass `","` for a standard comma-separated file or, for
  example, `";"` or `"\t"`. Only the first character is used; an empty string falls back to `,`.
- **`t`** — the [`Table`](tables-and-columns.md#tables) to export.
- **`params`** — [file writing options](file-options.md). The `.csv` extension is added
  automatically when `Extension` is empty.

## Basic example

```go
package main

import (
	"log"

	"github.com/Zapharaos/go-spit"
)

func main() {
	data := spit.DataSlice{
		{"name": "John Doe", "age": 30, "salary": 75000.50},
		{"name": "Jane Smith", "age": 28, "salary": 82000},
	}

	columns := spit.Columns{
		spit.NewColumn("name", "Full Name"),
		spit.NewColumn("age", "Age"),
		spit.NewColumn("salary", "Salary"),
	}

	table := spit.NewTable(data, columns, true)

	result, err := spit.ExportCSV(",", table, spit.FileWriteParams{
		Filename: "employees",
	})
	if err != nil {
		log.Fatal(err)
	}
	defer result.RemoveFile()

	log.Printf("created %s", result.Filepath)
}
```

Output (`employees.csv`):

```csv
Full Name,Age,Salary
John Doe,30,75000.5
Jane Smith,28,82000
```

## Headers

When the table is created with `writeHeader == true`, headers are generated from the column
labels. [Hierarchical columns](tables-and-columns.md#hierarchical-grouped-columns) produce one
header row per level: parent labels appear on the upper rows and leaf labels on the lower rows,
with empty cells used to visually span groups.

To omit headers entirely, build the table with `writeHeader == false`:

```go
table := spit.NewTable(data, columns, false)
```

## Formatting values

Use `Column.WithFormat` to format values. For dates, pass a Go time layout:

```go
columns := spit.Columns{
	spit.NewColumn("created_at", "Created At").WithFormat(time.RFC1123Z),
}
```

Strings are written as-is, and numbers/booleans use their default representation. See
[`FormatValue`](../reference/api.md) for the exact formatting rules.

## List (slice) values

If a value is a slice, set `Table.ListSeparator` to join its elements into a single cell:

```go
data := spit.DataSlice{
	{"tags": []interface{}{"go", "csv", "export"}},
}
columns := spit.Columns{spit.NewColumn("tags", "Tags")}

table := spit.NewTable(data, columns, true)
table.ListSeparator = "; "
```

The `tags` cell becomes `go; csv; export`.

!!! note "Styling and CSV"
    CSV is a plain-text format and does not support styles, borders or cell merging. Those
    options only affect [XLSX export](xlsx-export.md). Hierarchical headers, however, are fully
    supported in CSV.

## Image values

CSV cannot embed images. When a cell holds an [`Image`](tables-and-columns.md#images), CSV writes
its text fallback — the `URL` (or `AltText` when no URL is set):

```go
data := spit.DataSlice{
	{"company": "Acme", "logo": spit.NewImageURL("https://acme.com/logo.png")},
}
// The logo cell becomes: https://acme.com/logo.png
```
