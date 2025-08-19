// spreadsheet.go - spreadsheet abstraction and operations.
//
// This interface abstracts spreadsheet functionality, allowing for multiple implementations (e.g., Excel, Google Sheets) and libraries.
// It includes methods for file management, sheet management, column formatting, and table operations.

//go:generate mockgen -destination=spreadsheet_mock.go -package=spit . Spreadsheet

package spit

import (
	"io"
)

// Spreadsheet defines the interface for spreadsheet-specific operations.
type Spreadsheet interface {
	TableOperations // Embeds table-related operations (see TableOperations interface)

	// GetFile returns the underlying file object (implementation-specific).
	getFile() interface{}

	// CreateNewFile initializes a new spreadsheet file.
	createNewFile() error

	// SaveToWriter writes the spreadsheet to an io.Writer (e.g., file, buffer).
	saveToWriter(writer io.Writer) error

	// Close releases resources associated with the spreadsheet file.
	close() error

	// GetSheetName returns the current sheet name.
	getSheetName() string

	// SetSheetName sets the active sheet name.
	setSheetName(name string)

	// createSheet creates a new sheet with the current sheet name.
	createSheet() error

	// setActiveSheet sets the active sheet for subsequent operations.
	setActiveSheet() error

	// setColumnWidth sets the width of a column by its letter (e.g., "A", "B").
	setColumnWidth(colLetter string, width float64) error
}
