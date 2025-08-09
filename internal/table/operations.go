package table

// Operations defines table-specific operations interface
// This interface consolidates all operations that make sense for table processing
type Operations interface {
	GetTable() *Table

	GetCellValue(col, row int) (string, error)
	SetCellValue(col, row int, value interface{}) error

	MergeCells(startCol, startRow, endCol, endRow int) error
	IsCellMerged(col, row int) bool
	IsCellMergedHorizontally(col, row int) bool

	ApplyCellBorder(col, row int, side string, borderSide *BorderSide) error
	ApplyRangeBorder(startCol, startRow, endCol, endRow int, borderConfig BorderConfig) error
	HasExistingBorder(col, row int, side string) bool

	ApplyCellStyle(col, row int, style StyleConfig) error
	ApplyRangeStyle(startCol, startRow, endCol, endRow int, style StyleConfig) error

	GetColumnLetter(col int) string
	ProcessValue(value interface{}, format string) (interface{}, error)
}
