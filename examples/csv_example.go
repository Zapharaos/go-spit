// csv_example.go demonstrates how to use the CSV format
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
			"skills":     []interface{}{"Go", "Python", "Docker"},
		},
		{
			"name":       "Jane Smith",
			"age":        28,
			"email":      "jane@example.com",
			"salary":     82000.00,
			"department": "Marketing",
			"skills":     []interface{}{"Marketing", "Analytics", "SEO"},
		},
		{
			"name":       "Bob Johnson",
			"age":        35,
			"email":      "bob@example.com",
			"salary":     68000.75,
			"department": "Engineering",
			"skills":     []interface{}{"JavaScript", "React", "Node.js"},
		},
	}

	// Simple CSV with comma separator
	fmt.Println("Example 1: Simple CSV with comma separator")

	simpleColumns := go_spit.Columns{
		{Name: "name", Label: "Full Name"},
		{Name: "age", Label: "Age"},
		{Name: "email", Label: "Email Address"},
		{Name: "department", Label: "Department"},
		{Name: "salary", Label: "Salary"},
		{Name: "skills", Label: "Skills"},
	}

	table := &go_spit.Table{
		Data:          data,
		Columns:       simpleColumns,
		WriteHeader:   true,
		ListSeparator: "; ", // How to join array/slice values
	}

	// Create CSV instance with NewCsv
	csv := go_spit.NewCsv(",", table)

	// Write data to file
	options := go_spit.FileWriteOptions{
		Filename:      "employees",
		OverwriteFile: true,
	}

	result, err := csv.WriteDataToFile(options)
	if err != nil {
		log.Fatalf("Error writing CSV file: %v", err)
	}

	fmt.Printf("CSV file created: %s\n", result.FileName)
}
