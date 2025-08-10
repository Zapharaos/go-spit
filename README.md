[![PkgGoDev](https://pkg.go.dev/badge/mod/github.com/zapharaos/go-spit)](https://pkg.go.dev/mod/github.com/zapharaos/go-spit)
![Go Version](https://img.shields.io/badge/go%20version-%3E=1.24.1-61CFDD.svg?style=flat-square)
[![Go Report Card](https://goreportcard.com/badge/github.com/Zapharaos/go-spit)](https://goreportcard.com/report/github.com/Zapharaos/go-spit)
![GitHub License](https://img.shields.io/github/license/zapharaos/go-spit)

![GitHub Release](https://img.shields.io/github/v/release/zapharaos/go-spit)
![GitHub Actions Workflow Status](https://img.shields.io/github/actions/workflow/status/zapharaos/go-spit/golang.yml)
[![codecov](https://codecov.io/gh/Zapharaos/go-spit/graph/badge.svg?token=BL7YP0GTK9)](https://codecov.io/gh/Zapharaos/go-spit)

# go-spit

Go-spit is a Go package for flexible file exporting. It supports multiple formats and is designed for extensibility, making it easy to export data in various ways for reporting, data exchange, and automation.

## Features
- Export tabular data and spreadsheets
- Customizable file writing options (compression, overwrite, etc.)
- Multiple output formats

## Supported Formats
- CSV
- XLSX

## Installation

```sh
go get github.com/Zapharaos/go-spit
```

**Note:** Spit uses [Go Modules](https://go.dev/wiki/Modules) to manage dependencies.

## Usage Example

### CSV Export

Export tabular data to CSV with different separators:

```go
import "github.com/Zapharaos/go-spit"

// Create sample data
data := go_spit.DataSlice{
    {
        "name":       "John Doe",
        "age":        30,
        "email":      "john@example.com",
        "department": "Engineering",
        "skills":     []interface{}{"Go", "Python", "Docker"},
    },
    {
        "name":       "Jane Smith", 
        "age":        28,
        "email":      "jane@example.com",
        "department": "Marketing",
        "skills":     []interface{}{"Marketing", "Analytics", "SEO"},
    },
}

// Define columns
columns := go_spit.Columns{
    {Name: "name", Label: "Full Name"},
    {Name: "age", Label: "Age"},
    {Name: "email", Label: "Email Address"},
    {Name: "department", Label: "Department"},
    {Name: "skills", Label: "Skills"},
}

// Create table
table := &go_spit.Table{
    Data:          data,
    Columns:       columns,
    WriteHeader:   true,
    ListSeparator: "; ", // How to join array/slice values
}

// Create CSV with comma separator
csv := go_spit.NewCsv(",", table)

// Write to file
options := go_spit.FileWriteOptions{
    Filename:      "employees",
    OverwriteFile: true,
}

result, err := csv.WriteDataToFile(options)
if err != nil {
    log.Fatalf("Error writing CSV: %v", err)
}
```

### XLSX Export

Export data to Excel spreadsheets with hierarchical headers:

```go
import (
    "github.com/Zapharaos/go-spit"
    "github.com/xuri/excelize/v2"
)

// Create hierarchical columns
columns := go_spit.Columns{
    {
        Label: "Personal Info",
        Columns: go_spit.Columns{
            {Name: "name", Label: "Full Name"},
            {Name: "age", Label: "Age"},
            {Name: "email", Label: "Email Address"},
        },
    },
    {
        Label: "Work Info",
        Columns: go_spit.Columns{
            {Name: "department", Label: "Department"},
            {Name: "salary", Label: "Salary", Format: "currency"},
        },
    },
}

table := &go_spit.Table{
    Data:        data,
    Columns:     columns,
    WriteHeader: true,
}

// Method 1: Using NewXlsx with custom spreadsheet
file := excelize.NewFile()
spreadsheet := go_spit.NewSpreadsheetExcelize(file, "Employee Data", table)
xlsx := go_spit.NewXlsx(spreadsheet)

// Method 2: Using convenience function
xlsx2 := go_spit.NewXlsxWithExcelize(excelize.NewFile(), "Sales Report", table)

// Write to file
options := go_spit.FileWriteOptions{
    Filename:      "employee_report",
    OverwriteFile: true,
}

result, err := xlsx.WriteDataToFile(options)
if err != nil {
    log.Fatalf("Error writing XLSX: %v", err)
}
```

### Advanced File Options

Write files with compression and custom settings:

```go
options := go_spit.FileWriteOptions{
    Filename:      "report",
    UseGzip:       true,     // Enable compression
    OverwriteFile: true,     // Overwrite existing files
}

result, err := csv.WriteDataToFile(options)
```

### Complete Examples

For complete working examples with various use cases, see:
- [CSV Examples](examples/csv_example.go) - Demonstrates different separators, hierarchical headers, and compression
- [XLSX Examples](examples/xlsx_example.go) - Shows hierarchical columns, multiple sheets, and advanced formatting

## Development

Run the test suite:

```shell
make test
```

Run the test suite with coverage:

```shell
make coverage
```

Run linters:

```shell
make lint # pass -j option to run them in parallel
```

Some linter violations can automatically be fixed:

```shell
make fmt
```


## Contributing

We welcome contributions to the go-spit library! If you have a bug fix, feature request, or improvement, please open an issue or pull request on GitHub. We appreciate your help in making go-spit better for everyone. If you are interested in contributing to the go-spit library, please check out our [contributing guidelines](CONTRIBUTING.md) for more information on how to get started.

## License

The project is licensed under the [MIT License](LICENSE).
