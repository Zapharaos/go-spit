package gsheets

import (
	"fmt"
	"strconv"
	"strings"

	spit "github.com/Zapharaos/go-spit"
	"google.golang.org/api/sheets/v4"
)

// extendedValue converts a Go value to a Sheets ExtendedValue, preserving numeric and
// boolean types so Sheets treats them natively.
func extendedValue(value interface{}) *sheets.ExtendedValue {
	switch v := value.(type) {
	case nil:
		return nil
	case bool:
		b := v
		return &sheets.ExtendedValue{BoolValue: &b}
	case string:
		if v == "" {
			return nil
		}
		s := v
		return &sheets.ExtendedValue{StringValue: &s}
	case float64:
		return numberValue(v)
	case float32:
		return numberValue(float64(v))
	case int:
		return numberValue(float64(v))
	case int8:
		return numberValue(float64(v))
	case int16:
		return numberValue(float64(v))
	case int32:
		return numberValue(float64(v))
	case int64:
		return numberValue(float64(v))
	case uint:
		return numberValue(float64(v))
	case uint8:
		return numberValue(float64(v))
	case uint16:
		return numberValue(float64(v))
	case uint32:
		return numberValue(float64(v))
	case uint64:
		return numberValue(float64(v))
	default:
		s := fmt.Sprintf("%v", v)
		return &sheets.ExtendedValue{StringValue: &s}
	}
}

func numberValue(f float64) *sheets.ExtendedValue {
	return &sheets.ExtendedValue{NumberValue: &f, ForceSendFields: []string{"NumberValue"}}
}

// isNumeric reports whether a value is one of Go's numeric types.
func isNumeric(value interface{}) bool {
	switch value.(type) {
	case int, int8, int16, int32, int64,
		uint, uint8, uint16, uint32, uint64,
		float32, float64:
		return true
	}
	return false
}

// applyStyle overlays a spit.Style onto a Sheets CellFormat, preserving previously set
// properties (e.g. borders).
func applyStyle(cf *sheets.CellFormat, s spit.Style) {
	if s.Bold || s.Italic || s.FontSize > 0 || s.FontFamily != "" || s.TextColor != "" {
		if cf.TextFormat == nil {
			cf.TextFormat = &sheets.TextFormat{}
		}
		tf := cf.TextFormat
		if s.Bold {
			tf.Bold = true
		}
		if s.Italic {
			tf.Italic = true
		}
		if s.FontSize > 0 {
			tf.FontSize = int64(s.FontSize)
		}
		if s.FontFamily != "" {
			tf.FontFamily = s.FontFamily
		}
		if s.TextColor != "" {
			tf.ForegroundColor = hexColor(s.TextColor)
		}
	}
	if s.BackgroundColor != "" {
		cf.BackgroundColor = hexColor(s.BackgroundColor)
	}
	if s.Alignment != spit.AlignmentNone {
		horizontal, vertical := s.Alignment.GetAlignmentValues()
		cf.HorizontalAlignment = strings.ToUpper(horizontal)
		cf.VerticalAlignment = verticalAlignment(vertical)
	}
	if s.NumFmt != "" {
		cf.NumberFormat = &sheets.NumberFormat{Type: "NUMBER", Pattern: s.NumFmt}
	}
}

// verticalAlignment maps an internal vertical token to a Sheets vertical alignment.
func verticalAlignment(v string) string {
	switch v {
	case "center":
		return "MIDDLE"
	case "bottom":
		return "BOTTOM"
	default:
		return "TOP"
	}
}

// borderStyle maps a spit.BorderStyle to a Sheets border style token.
func borderStyle(bs spit.BorderStyle) string {
	switch bs {
	case spit.BorderStyleMedium:
		return "SOLID_MEDIUM"
	case spit.BorderStyleThick:
		return "SOLID_THICK"
	case spit.BorderStyleDashed:
		return "DASHED"
	case spit.BorderStyleDotted:
		return "DOTTED"
	case spit.BorderStyleDouble:
		return "DOUBLE"
	default:
		return "SOLID"
	}
}

// hexColor parses a hex color ("#RRGGBB" or "RRGGBB") into a Sheets Color, or nil.
func hexColor(s string) *sheets.Color {
	s = strings.TrimPrefix(strings.TrimSpace(s), "#")
	if len(s) != 6 {
		return nil
	}
	r, err1 := strconv.ParseInt(s[0:2], 16, 0)
	green, err2 := strconv.ParseInt(s[2:4], 16, 0)
	b, err3 := strconv.ParseInt(s[4:6], 16, 0)
	if err1 != nil || err2 != nil || err3 != nil {
		return nil
	}
	return &sheets.Color{
		Red:             float64(r) / 255,
		Green:           float64(green) / 255,
		Blue:            float64(b) / 255,
		Alpha:           1,
		ForceSendFields: []string{"Red", "Green", "Blue", "Alpha"},
	}
}

// blackColor returns an opaque black Sheets color for borders.
func blackColor() *sheets.Color {
	return &sheets.Color{Alpha: 1, ForceSendFields: []string{"Red", "Green", "Blue", "Alpha"}}
}

// toImage extracts a spit.Image from a cell value (accepts Image and *Image).
func toImage(value interface{}) (spit.Image, bool) {
	switch v := value.(type) {
	case spit.Image:
		return v, true
	case *spit.Image:
		if v != nil {
			return *v, true
		}
	}
	return spit.Image{}, false
}

// columnLetter returns the spreadsheet column letters for a 1-based index.
func columnLetter(col int) string {
	if col <= 0 {
		return ""
	}
	var b []byte
	for col > 0 {
		col--
		b = append([]byte{byte('A' + col%26)}, b...)
		col /= 26
	}
	return string(b)
}
