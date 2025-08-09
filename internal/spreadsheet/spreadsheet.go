package spreadsheet

import (
	"io"

	"github.com/Zapharaos/go-spit/internal/table"
)

// Spreadsheet defines the interface for spreadsheet-specific operations
type Spreadsheet interface {
	table.Operations

	GetTable() *table.Table

	GetFile() interface{} // Returns the underlying file object
	CreateNewFile() error
	SaveToWriter(writer io.Writer) error
	Close() error

	GetSheetName() string
	SetSheetName(name string)
	CreateSheet() error
	SetActiveSheet() error

	SetColumnWidth(colLetter string, width float64) error
}
