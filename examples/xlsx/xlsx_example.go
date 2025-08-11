// xlsx_example.go demonstrates how to use the XLSX format
package main

import (
	"fmt"
	"log"

	"github.com/Zapharaos/go-spit"
)

func main() {
	// Create sample data
	data := go_spit.DataSlice{
		{
			"name":       "John Doe",
			"age":        30,
			"email":      "john@example.com",
			"salary":     75000.50,
			"department": "Engineering",
		},
		{
			"name":       "Jane Smith",
			"age":        28,
			"email":      "jane@example.com",
			"salary":     82000.00,
			"department": "Marketing",
		},
		{
			"name":       "Bob Johnson",
			"age":        35,
			"email":      "bob@example.com",
			"salary":     68000.75,
			"department": "Engineering",
		},
	}

	// Define columns with hierarchical structure
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

	// Create a table with the data and columns
	table := &go_spit.Table{
		Data:        data,
		Columns:     columns,
		WriteHeader: true,
		Limit:       0, // No limit
	}

	// Method 1 : Create XLSX instance with NewXlsx
	// We can use nil file to let the library create a new file
	spreadsheet := go_spit.NewSpreadsheetExcelize(nil, "Employee Data", table)
	xlsx := go_spit.NewXlsx(spreadsheet)

	options := go_spit.FileWriteOptions{
		Filename:      "employee_report",
		UseTempFile:   true,
		OverwriteFile: true,
	}

	result, err := xlsx.WriteDataToFile(options)
	if err != nil {
		log.Fatalf("Error writing XLSX file: %v", err)
	}

	// Method 2 : Create XLSX instance with NewXlsxWithExcelize convenience function
	// We can reuse the file created earlier and write to a new sheet
	xlsx = go_spit.NewXlsxWithExcelize(spreadsheet.File, "Sales Report", table)

	result, err = xlsx.WriteDataToFile(options)
	if err != nil {
		log.Fatalf("Error writing XLSX file: %v", err)
	}

	fmt.Printf("XLSX file created: %s\n", result.FileName)
}
