package table

import (
	"fmt"
	"time"

	"github.com/Zapharaos/go-spit/internal/utils"
	"github.com/xuri/excelize/v2"
)

// Excelize provides Excelize-specific operations for table handling
type Excelize struct {
	File      *excelize.File
	SheetName string
	Table     *Table
}

func NewExcelize(file *excelize.File, sheetName string, table *Table) *Excelize {
	return &Excelize{
		File:      file,
		SheetName: sheetName,
		Table:     table,
	}
}

func (e *Excelize) GetTable() *Table {
	return e.Table
}

func (e *Excelize) GetCellValue(col, row int) (string, error) {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return "", err
	}
	return e.File.GetCellValue(e.SheetName, cellRef)
}

func (e *Excelize) SetCellValue(col, row int, value interface{}) error {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return err
	}
	return e.File.SetCellValue(e.SheetName, cellRef, value)
}

func (e *Excelize) MergeCells(startCol, startRow, endCol, endRow int) error {
	startCell, err1 := excelize.CoordinatesToCellName(startCol, startRow)
	endCell, err2 := excelize.CoordinatesToCellName(endCol, endRow)
	if err1 != nil || err2 != nil {
		return fmt.Errorf("failed to convert coordinates: %v, %v", err1, err2)
	}
	return e.File.MergeCell(e.SheetName, startCell, endCell)
}

func (e *Excelize) IsCellMerged(col, row int) bool {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return false
	}
	return e.isCellMerged(cellRef)
}

func (e *Excelize) IsCellMergedHorizontally(col, row int) bool {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return false
	}
	return e.isCellMergedHorizontally(cellRef)
}

func (e *Excelize) ApplyCellBorder(col, row int, side string, borderSide *BorderSide) error {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return err
	}
	return e.applyCellBorder(cellRef, side, borderSide)
}

func (e *Excelize) ApplyRangeBorder(startCol, startRow, endCol, endRow int, borderConfig BorderConfig) error {
	for row := startRow; row <= endRow; row++ {
		for col := startCol; col <= endCol; col++ {
			if col == startCol && borderConfig.Left != nil {
				if err := e.ApplyCellBorder(col, row, "left", borderConfig.Left); err != nil {
					return err
				}
			}
			if col == endCol && borderConfig.Right != nil {
				if err := e.ApplyCellBorder(col, row, "right", borderConfig.Right); err != nil {
					return err
				}
			}
			if row == startRow && borderConfig.Top != nil {
				if err := e.ApplyCellBorder(col, row, "top", borderConfig.Top); err != nil {
					return err
				}
			}
			if row == endRow && borderConfig.Bottom != nil {
				if err := e.ApplyCellBorder(col, row, "bottom", borderConfig.Bottom); err != nil {
					return err
				}
			}
		}
	}
	return nil
}

func (e *Excelize) HasExistingBorder(col, row int, side string) bool {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return false
	}
	styleID, err := e.File.GetCellStyle(e.SheetName, cellRef)
	if err != nil {
		return false
	}
	// Check if there's a style applied (simple check)
	return styleID > 0
}

func (e *Excelize) ApplyCellStyle(col, row int, style StyleConfig) error {
	cellRef, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return err
	}
	return e.applyCellStyle(cellRef, style)
}

func (e *Excelize) ApplyRangeStyle(startCol, startRow, endCol, endRow int, style StyleConfig) error {
	startCell, err1 := excelize.CoordinatesToCellName(startCol, startRow)
	endCell, err2 := excelize.CoordinatesToCellName(endCol, endRow)
	if err1 != nil || err2 != nil {
		return fmt.Errorf("failed to convert coordinates: %v, %v", err1, err2)
	}

	excelStyle := convertStyleToExcelizeStyle(style)
	styleID, err := e.File.NewStyle(excelStyle)
	if err != nil {
		return err
	}

	return e.File.SetCellStyle(e.SheetName, startCell, endCell, styleID)
}

func (e *Excelize) GetColumnLetter(col int) string {
	letter, _ := excelize.ColumnNumberToName(col)
	return letter
}

func (e *Excelize) ProcessValue(value interface{}, format string) (interface{}, error) {
	switch v := value.(type) {
	case []interface{}:
		if e.Table.ListSeparator != "" {
			return utils.ConvertSliceToString(v, format, e.Table.ListSeparator)
		}
		return fmt.Sprintf("%v", v), nil
	case time.Time:
		if format != "" {
			return v.Format(format), nil
		}
		return v, nil
	case *time.Time:
		if v != nil {
			if format != "" {
				return v.Format(format), nil
			}
			return *v, nil
		}
		return "", nil
	case string:
		// Skip formatting for string values, even if format is specified
		// This prevents format conflicts (e.g., "Total" being formatted as date)
		return v, nil
	case int, int8, int16, int32, int64:
		return v, nil
	case uint, uint8, uint16, uint32, uint64:
		return v, nil
	case float32, float64:
		return v, nil
	case bool:
		return v, nil
	default:
		if format != "" {
			var err error
			value, err = utils.FormatValue(value, format)
			if err != nil {
				return "", err
			}
		}
		return fmt.Sprintf("%v", value), nil
	}
}

func (e *Excelize) isCellMerged(cellRef string) bool {
	mergedCells, err := e.File.GetMergeCells(e.SheetName)
	if err != nil {
		return false
	}
	for _, mergeCell := range mergedCells {
		if isCellInRange(cellRef, mergeCell.GetStartAxis(), mergeCell.GetEndAxis()) {
			return true
		}
	}
	return false
}

func (e *Excelize) isCellMergedHorizontally(cellRef string) bool {
	mergedCells, err := e.File.GetMergeCells(e.SheetName)
	if err != nil {
		return false
	}
	for _, mergeCell := range mergedCells {
		if isCellInRange(cellRef, mergeCell.GetStartAxis(), mergeCell.GetEndAxis()) {
			startCol, startRow, _ := excelize.CellNameToCoordinates(mergeCell.GetStartAxis())
			endCol, endRow, _ := excelize.CellNameToCoordinates(mergeCell.GetEndAxis())
			return startRow == endRow && startCol != endCol
		}
	}
	return false
}

func (e *Excelize) applyCellBorder(cellRef, borderType string, borderSide *BorderSide) error {
	if borderSide == nil || borderSide.Style == BorderStyleNone {
		return nil
	}

	excelBorderStyle := convertBorderStyleToExcelize(borderSide.Style)

	// Create the style with border
	style := &excelize.Style{}

	switch borderType {
	case "left":
		style.Border = []excelize.Border{{Type: "left", Color: "000000", Style: excelBorderStyle}}
	case "right":
		style.Border = []excelize.Border{{Type: "right", Color: "000000", Style: excelBorderStyle}}
	case "top":
		style.Border = []excelize.Border{{Type: "top", Color: "000000", Style: excelBorderStyle}}
	case "bottom":
		style.Border = []excelize.Border{{Type: "bottom", Color: "000000", Style: excelBorderStyle}}
	default:
		return fmt.Errorf("unsupported border type: %s", borderType)
	}

	styleID, err := e.File.NewStyle(style)
	if err != nil {
		return err
	}

	return e.File.SetCellStyle(e.SheetName, cellRef, cellRef, styleID)
}

func (e *Excelize) applyCellStyle(cellRef string, style StyleConfig) error {
	excelStyle := convertStyleToExcelizeStyle(style)
	styleID, err := e.File.NewStyle(excelStyle)
	if err != nil {
		return err
	}
	return e.File.SetCellStyle(e.SheetName, cellRef, cellRef, styleID)
}

func convertBorderStyleToExcelize(style BorderStyle) int {
	switch style {
	case BorderStyleNone:
		return 0
	case BorderStyleThin:
		return 1
	case BorderStyleMedium:
		return 2
	case BorderStyleDashed:
		return 3
	case BorderStyleDotted:
		return 4
	case BorderStyleThick:
		return 5
	case BorderStyleDouble:
		return 6
	default:
		return 1
	}
}

func convertStyleToExcelizeStyle(style StyleConfig) *excelize.Style {
	excelStyle := &excelize.Style{}

	if style.Bold || style.Italic || style.FontSize > 0 || style.FontFamily != "" || style.TextColor != "" {
		font := &excelize.Font{}
		if style.Bold {
			font.Bold = true
		}
		if style.Italic {
			font.Italic = true
		}
		if style.FontSize > 0 {
			font.Size = style.FontSize
		}
		if style.FontFamily != "" {
			font.Family = style.FontFamily
		}
		if style.TextColor != "" {
			font.Color = style.TextColor
		}
		if style.Underline != "" {
			font.Underline = style.Underline
		}
		excelStyle.Font = font
	}

	if style.BackgroundColor != "" {
		excelStyle.Fill = excelize.Fill{
			Type:    "pattern",
			Color:   []string{style.BackgroundColor},
			Pattern: 1,
		}
	}

	// TODO : check function
	if style.Alignment != AlignmentNone {
		alignment := &excelize.Alignment{}
		switch style.Alignment {
		case AlignmentLeft:
			alignment.Horizontal = "left"
		case AlignmentCenter:
			alignment.Horizontal = "center"
		case AlignmentRight:
			alignment.Horizontal = "right"
		}
		excelStyle.Alignment = alignment
	}

	return excelStyle
}

func isCellInRange(cellRef, startRef, endRef string) bool {
	col, row, err := excelize.CellNameToCoordinates(cellRef)
	if err != nil {
		return false
	}
	startCol, startRow, err := excelize.CellNameToCoordinates(startRef)
	if err != nil {
		return false
	}
	endCol, endRow, err := excelize.CellNameToCoordinates(endRef)
	if err != nil {
		return false
	}
	return col >= startCol && col <= endCol && row >= startRow && row <= endRow
}
