# Quick Start

This guide walks you through the core workflow of go-spit: define your **data** and
**columns**, build a **table**, and export it to a file.

The export workflow is always the same three steps:

1. Describe the rows with a [`DataSlice`](../user-guide/tables-and-columns.md#data).
2. Describe the headers with [`Columns`](../user-guide/tables-and-columns.md#columns).
3. Build a [`Table`](../user-guide/tables-and-columns.md#tables) and call an exporter
   (`ExportCSV` or `ExportXLSX`).

## Export a CSV file

```go
package main

import (
	"log"

	"github.com/Zapharaos/go-spit"
)

func main() {
	// 1. Data: a slice of maps, one map per row.
	data := spit.DataSlice{
		{"name": "John Doe", "age": 30, "salary": 75000.50},
		{"name": "Jane Smith", "age": 28, "salary": 82000},
	}

	// 2. Columns: map data keys to header labels.
	columns := spit.Columns{
		spit.NewColumn("name", "Full Name"),
		spit.NewColumn("age", "Age"),
		spit.NewColumn("salary", "Salary"),
	}

	// 3. Table: the third argument enables header generation.
	table := spit.NewTable(data, columns, true)

	// Export to CSV using "," as the field separator.
	result, err := spit.ExportCSV(",", table, spit.FileWriteParams{
		Filename: "employees",
	})
	if err != nil {
		log.Fatal(err)
	}

	log.Printf("created %s", result.Filepath)

	// Remove the file when you no longer need it.
	defer func() {
		if err := result.RemoveFile(); err != nil {
			log.Printf("failed to remove file: %v", err)
		}
	}()
}
```

This produces an `employees.csv` file in the current directory:

```csv
Full Name,Age,Salary
John Doe,30,75000.5
Jane Smith,28,82000
```

## Export an XLSX file

XLSX export uses a [`Spreadsheet`](../user-guide/xlsx-export.md) implementation. go-spit ships
with an [Excelize](https://github.com/xuri/excelize)-backed implementation,
`NewSpreadsheetExcelize`.

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

	// Hierarchical columns: "Details" groups two sub-columns.
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

## Next steps

- Understand the data model in [Tables, Data & Columns](../user-guide/tables-and-columns.md).
- Customize output with [Styling, Borders & Merging](../user-guide/styling.md).
- Control where and how files are written with [File Options](../user-guide/file-options.md).
- Browse complete, runnable programs in the
  [`examples/`](https://github.com/Zapharaos/go-spit/tree/main/examples) directory.
