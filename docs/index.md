# go-spit

[![PkgGoDev](https://pkg.go.dev/badge/mod/github.com/Zapharaos/go-spit)](https://pkg.go.dev/mod/github.com/Zapharaos/go-spit)
![Go Version](https://img.shields.io/badge/go%20version-%3E=1.24.1-61CFDD.svg?style=flat-square)
[![Go Report Card](https://goreportcard.com/badge/github.com/Zapharaos/go-spit)](https://goreportcard.com/report/github.com/Zapharaos/go-spit)
![GitHub License](https://img.shields.io/github/license/Zapharaos/go-spit)

**go-spit** is a Go package for flexible file exporting. It supports multiple formats and is
designed for extensibility, making it easy to export data in various ways for reporting, data
exchange and automation.

## Features

- Export tabular data and spreadsheets
- Customizable file writing options (compression, overwrite, temporary files, etc.)
- Multiple output formats
- Hierarchical (grouped) column headers
- Advanced XLSX styling: fonts, colors, alignment, borders and cell merging
- Pluggable logging that adapts to your existing logger

## Supported Formats

| Format   | Description                                                                 |
|----------|-----------------------------------------------------------------------------|
| **CSV**  | Simple tabular data with custom delimiters and multi-level headers          |
| **XLSX** | Advanced spreadsheets with styling, borders, merging and hierarchical headers |

## Where to Go Next

<div class="grid cards" markdown>

- :material-download: **[Installation](getting-started/installation.md)**

    Add go-spit to your Go module.

- :material-rocket-launch: **[Quick Start](getting-started/quickstart.md)**

    Export your first CSV and XLSX files in minutes.

- :material-book-open-variant: **[User Guide](user-guide/tables-and-columns.md)**

    Learn tables, columns, styling, merging and file options in depth.

- :material-code-tags: **[API Reference](reference/api.md)**

    Full Go API documentation on pkg.go.dev.

</div>

## A Quick Taste

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
}
```

## License

The project is licensed under the [MIT License](https://github.com/Zapharaos/go-spit/blob/main/LICENSE).
