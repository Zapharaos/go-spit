[![PkgGoDev](https://pkg.go.dev/badge/mod/github.com/Zapharaos/go-spit)](https://pkg.go.dev/mod/github.com/Zapharaos/go-spit)
![Go Version](https://img.shields.io/badge/go%20version-%3E=1.24.1-61CFDD.svg?style=flat-square)
[![Go Report Card](https://goreportcard.com/badge/github.com/Zapharaos/go-spit)](https://goreportcard.com/report/github.com/Zapharaos/go-spit)
![GitHub License](https://img.shields.io/github/license/Zapharaos/go-spit)

![GitHub Release](https://img.shields.io/github/v/release/Zapharaos/go-spit)
![GitHub Actions Workflow Status](https://img.shields.io/github/actions/workflow/status/Zapharaos/go-spit/golang.yml)
[![codecov](https://codecov.io/gh/Zapharaos/go-spit/graph/badge.svg?token=VNQGKOP6ZX)](https://codecov.io/gh/Zapharaos/go-spit)

# go-spit

Go-spit is a Go package for flexible file exporting. It supports multiple formats and is designed for extensibility, making it easy to export data in various ways for reporting, data exchange and automation.

## Features
- Export tabular data and spreadsheets
- Customizable file writing options (compression, overwrite, temporary, etc.)
- Multiple output formats

## Supported Formats
- **CSV**: Simple tabular data with custom delimiters
- **XLSX**: Advanced spreadsheets with styling, borders, merging, and hierarchical headers

## Installation

```sh
go get github.com/Zapharaos/go-spit
```

**Note:** Spit uses [Go Modules](https://go.dev/wiki/Modules) to manage dependencies.

## Usage Examples

### CSV Export

```go
package main

import (
    "log"
    "github.com/Zapharaos/go-spit"
)

func main() {
    // Sample data
    data := spit.DataSlice{
        {"name": "John Doe", "age": 30, "salary": 75000.50},
        {"name": "Jane Smith", "age": 28, "salary": 82000},
    }

    // Define columns with formatting
    columns := spit.Columns{
        spit.NewColumn("name", "Full Name"),
        spit.NewColumn("age", "Age"),
        spit.NewColumn("salary", "Salary"),
    }

    // Create table using constructor and fluent API
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

### XLSX Export

```go
package main

import (
    "log"
    "github.com/Zapharaos/go-spit"
)

func main() {
    // Sample data
    data := spit.DataSlice{
        {"name": "John Doe", "age": 30, "department": "Engineering"},
        {"name": "Jane Smith", "age": 28, "department": "Marketing"},
    }

    // Define hierarchical columns
    columns := spit.Columns{
        spit.NewColumn("name", "Employee Name"),
        spit.NewColumn("", "Details").
            WithSubColumns(spit.Columns{
                spit.NewColumn("age", "Age"),
                spit.NewColumn("department", "Department"),
            }),
    }

    // Create table with row and cell options
    table := spit.NewTable(data, columns, true)
    spreadsheet := spit.NewSpreadsheetExcelize("Employee Report", table)
    result, err := spit.ExportXLSX(spreadsheet, spit.FileWriteParams{
        Filename: "advanced_report",
    })
    if err != nil {
        log.Fatal(err)
    }
    defer result.RemoveFile()
}
```

#### XLSX Advanced Features

The XLSX format supports advanced styling and formatting options:

- **Styles**: Font family, size, colors, bold, italic, alignment
- **Borders**: Thin, medium, thick, double, dashed borders with inner border support
- **Cell Merging**: Vertical and horizontal merging based on identical or empty values
- **Row Options**: Apply styling, borders, merging options to entire rows
- **Cell Options**: Fine-grained styling, borders, merging options for individual cells
- **Column Formatting**: Currency, date, custom number formats

### File Options

Write files with compression and custom settings:

```go
options := go_spit.FileWriteOptions{
    Filename:      "report", // Without extension, which will be added based on format
    Filepath:      "/path/to/directory", // Could be "." or empty as well
    UseTempFile:   false,    // Enable dedicated temporary files
    UseGzip:       true,     // Enable compression
    OverwriteFile: true,     // Overwrite existing file
}
```

## Development

Install dependencies:
```shell
make dev-deps
```

Run unit tests and generate coverage report:
```shell
make test-unit
```

Run linters:
```shell
make lint
```

Some linter violations can automatically be fixed:
```shell
make fmt
```

## Contributing

We welcome contributions to the go-spit library! If you have a bug fix, feature request, or improvement, please open an issue or pull request on GitHub. We appreciate your help in making go-spit better for everyone. If you are interested in contributing to the go-spit library, please check out our [contributing guidelines](CONTRIBUTING.md) for more information on how to get started.

## License

The project is licensed under the [MIT License](LICENSE).
