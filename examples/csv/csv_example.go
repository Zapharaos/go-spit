// csv_example.go demonstrates how to use the CSV format

package main

import (
	"log"
	"math/rand"
	"time"

	"github.com/Zapharaos/go-spit"
)

func main() {
	// Define columns using the new fluent API constructors and setters
	columns := spit.Columns{
		spit.NewColumn("name", "Full Name"),
		spit.NewColumn("age", "Age"),
		spit.NewColumn("email", "Email Address"),
		spit.NewColumn("department", "Department"),
		spit.NewColumn("salary", "Salary"),
		spit.NewColumn("created_at", "Created At").WithFormat(time.RFC1123Z),
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

	// Create a table using the new constructor and fluent API
	table := spit.NewTable(data, columns, true)

	// File parameters for writing CSV
	params := spit.FileWriteParams{
		Filename: "employees",
	}

	// Create CSV instance with NewCsv
	result, err := spit.ExportCSV(",", table, params)
	if err != nil {
		log.Fatalf("Error writing CSV file: %v", err)
	}

	log.Printf("Successfully created XLSX file: %s", result.Filepath)

	// Uncomment to remove the file after creation
	// defer func() {
	//     if closeErr := result.RemoveFile(); closeErr != nil {
	//         log.Printf("Failed to remove XLSX file: %v", closeErr)
	//     }
	// }()
}
