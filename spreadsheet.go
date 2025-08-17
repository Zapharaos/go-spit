// spreadsheet.go - spreadsheet abstraction and operations.
//
// This interface abstracts spreadsheet functionality, allowing for multiple implementations (e.g., Excel, Google Sheets) and libraries.
// It includes methods for file management, sheet management, column formatting, and table operations.

//go:generate mockgen -destination=mocks/spreadsheet_mock.go -package=mocks . Spreadsheet

package spit

import (
	"io"
)

// Spreadsheet defines the interface for spreadsheet-specific operations.
type Spreadsheet interface {
	tableOperations // Embeds table-related operations (see tableOperations interface)

	// getFile returns the underlying file object (implementation-specific).
	getFile() interface{}

	// createNewFile initializes a new spreadsheet file.
	createNewFile() error

	// saveToWriter writes the spreadsheet to an io.Writer (e.g., file, buffer).
	saveToWriter(writer io.Writer) error

	// close releases resources associated with the spreadsheet file.
	close() error

	// getSheetName returns the current sheet name.
	getSheetName() string

	// setSheetName sets the active sheet name.
	setSheetName(name string)

	// createSheet creates a new sheet with the current sheet name.
	createSheet() error

	// setActiveSheet sets the active sheet for subsequent operations.
	setActiveSheet() error

	// setColumnWidth sets the width of a column by its letter (e.g., "A", "B").
	setColumnWidth(colLetter string, width float64) error
}
