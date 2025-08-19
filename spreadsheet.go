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
	TableOperations // Embeds table-related operations (see TableOperations interface)

	// GetFile returns the underlying file object (implementation-specific).
	GetFile() interface{}

	// CreateNewFile initializes a new spreadsheet file.
	CreateNewFile() error

	// SaveToWriter writes the spreadsheet to an io.Writer (e.g., file, buffer).
	SaveToWriter(writer io.Writer) error

	// Close releases resources associated with the spreadsheet file.
	Close() error

	// GetSheetName returns the current sheet name.
	GetSheetName() string

	// SetSheetName sets the active sheet name.
	SetSheetName(name string)

	// CreateSheet creates a new sheet with the current sheet name.
	CreateSheet() error

	// SetActiveSheet sets the active sheet for subsequent operations.
	SetActiveSheet() error

	// SetColumnWidth sets the width of a column by its letter (e.g., "A", "B").
	SetColumnWidth(colLetter string, width float64) error
}
