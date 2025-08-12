// csv_example.go demonstrates how to use the CSV format

package main

import (
	"log"
	"math/rand"
	"time"

	"github.com/Zapharaos/go-spit"
)

func main() {
	// Define columns
	columns := []spit.Column{
		{Name: "name", Label: "Full Name"},
		{Name: "age", Label: "Age"},
		{Name: "email", Label: "Email Address"},
		{Name: "department", Label: "Department"},
		{Name: "salary", Label: "Salary", Format: "$%.2f"},
		{Name: "created_at", Label: "Created At", Format: time.RFC1123Z},
	}

	// Create sample data
	data := spit.DataSlice{
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
			"salary":     82000,
			"department": "Marketing",
		},
		{
			"name":       "Bob Johnson",
			"age":        35,
			"email":      "bob@example.com",
			"salary":     68000.7,
			"department": "Engineering",
		},
	}

	// Add a random created_at field to each data row
	for i := range data {
		data[i]["created_at"] = time.Unix(rand.Int63n(time.Now().Unix()), 0).UTC()
	}

	// Create a table with the data and columns
	table := &spit.Table{
		Data:        data,
		Columns:     columns,
		WriteHeader: true,
	}

	// File parameters for writing CSV
	params := spit.FileWriteParams{
		Filename:    "employees",
		UseTempFile: true,
	}

	// Create CSV instance with NewCsv
	result, err := spit.ExportCSV(",", table, params)
	if err != nil {
		log.Fatalf("Error writing CSV file: %v", err)
	}

	defer func() {
		if closeErr := result.RemoveFile(); closeErr != nil {
			log.Printf("Failed to remove CSV file: %v", closeErr)
		}
	}()
}
