package table

// RowConfigs maps row indices to their specific configurations.
type RowConfigs map[int]RowConfig

// RowConfig represents configuration options for a specific row in the export.
// This allows fine-grained control over individual rows, overriding default
// column-based settings when needed.
type RowConfig struct {
	RowIndex int           // The 0-based index of the row this configuration applies to
	Border   *BorderConfig // Optional border configuration for the entire row
	Merge    *MergeConfig  // Optional merge configuration that overrides column settings
	Style    *StyleConfig  // Optional style configuration for the entire row
}

// CellConfigs provides cell-level configuration mapping.
// The outer map keys are column indices, inner map keys are row indices.
type CellConfigs map[int]CellConfigsByRow

// CellConfigsByRow maps row indices to cell configurations for a specific column.
type CellConfigsByRow map[int]CellConfig

// CellConfig represents configuration options for a specific cell.
// This provides the finest level of control, allowing individual cells
// to override both column and row settings.
type CellConfig struct {
	RowIndex  int           // The 0-based row index of this cell
	ColIndex  int           // The 0-based column index of this cell
	Border    *BorderConfig // Optional border configuration for this cell
	Style     *StyleConfig  // Optional style configuration for this cell
	Mergeable bool          // Whether this cell can participate in merge operations
}

// BorderStyle represents the visual style of cell borders.
// These constants correspond to common border styles available in spreadsheet applications.
type BorderStyle int

const (
	BorderStyleNone   BorderStyle = 0 // No border
	BorderStyleThin   BorderStyle = 1 // Thin solid line
	BorderStyleMedium BorderStyle = 2 // Medium thickness solid line
	BorderStyleDashed BorderStyle = 3 // Dashed line
	BorderStyleDotted BorderStyle = 4 // Dotted line
	BorderStyleThick  BorderStyle = 5 // Thick solid line
	BorderStyleDouble BorderStyle = 6 // Double line
)

// BorderSide represents the configuration for one side of a cell border.
type BorderSide struct {
	Style BorderStyle // The visual style to apply to this border side
}

// BorderConfig represents complete border configuration for a cell, row, or column.
type BorderConfig struct {
	Left   *BorderSide   // Left border configuration
	Right  *BorderSide   // Right border configuration
	Top    *BorderSide   // Top border configuration
	Bottom *BorderSide   // Bottom border configuration
	Inner  *BorderConfig // Inner borders for ranges (used in some contexts)
}

// hasBorders checks if any borders are configured in this BorderConfig.
func (bc BorderConfig) hasBorders() bool {
	return (bc.Left != nil && bc.Left.Style != BorderStyleNone) ||
		(bc.Right != nil && bc.Right.Style != BorderStyleNone) ||
		(bc.Top != nil && bc.Top.Style != BorderStyleNone) ||
		(bc.Bottom != nil && bc.Bottom.Style != BorderStyleNone)
}

// SetInner creates inner border configuration with the same style for all sides.
func (bc BorderConfig) SetInner(style BorderStyle) BorderConfig {
	side := &BorderSide{Style: style}
	bc.Inner = &BorderConfig{
		Left:   side,
		Right:  side,
		Top:    side,
		Bottom: side,
	}
	return bc
}

// NewBorderConfig creates a BorderConfig with the same style applied to all sides.
func NewBorderConfig(style BorderStyle) BorderConfig {
	side := &BorderSide{Style: style}
	return BorderConfig{
		Left:   side,
		Right:  side,
		Top:    side,
		Bottom: side,
	}
}

// StyleConfig represents comprehensive styling configuration for cells.
type StyleConfig struct {
	Bold            bool          // Whether text should be bold
	Italic          bool          // Whether text should be italic
	Underline       string        // Underline style (format-specific values)
	TextColor       string        // Text color (usually hex format: "#RRGGBB")
	BackgroundColor string        // Cell background color (usually hex format: "#RRGGBB")
	FontSize        float64       // Font size in points
	FontFamily      string        // Font family name (e.g., "Arial", "Times New Roman")
	Alignment       CellAlignment // Text alignment within the cell
}

// CellAlignment represents the alignment options for cell content.
type CellAlignment int

const (
	AlignmentNone         CellAlignment = iota // Use default alignment
	AlignmentLeft                              // Left-aligned text, top-aligned vertically
	AlignmentCenter                            // Center-aligned text, top-aligned vertically
	AlignmentRight                             // Right-aligned text, top-aligned vertically
	AlignmentTop                               // Left-aligned text, top-aligned vertically
	AlignmentMiddle                            // Left-aligned text, center-aligned vertically
	AlignmentBottom                            // Left-aligned text, bottom-aligned vertically
	AlignmentCenterMiddle                      // Center-aligned text, center-aligned vertically
	AlignmentLeftMiddle                        // Left-aligned text, center-aligned vertically
	AlignmentRightMiddle                       // Right-aligned text, center-aligned vertically
)

// getAlignmentValues converts CellAlignment enum to horizontal and vertical alignment strings.
func (ca CellAlignment) getAlignmentValues() (horizontal, vertical string) {
	switch ca {
	case AlignmentLeft:
		return "left", "top"
	case AlignmentCenter:
		return "center", "top"
	case AlignmentRight:
		return "right", "top"
	case AlignmentTop:
		return "left", "top"
	case AlignmentMiddle:
		return "left", "center"
	case AlignmentBottom:
		return "left", "bottom"
	case AlignmentCenterMiddle:
		return "center", "center"
	case AlignmentLeftMiddle:
		return "left", "center"
	case AlignmentRightMiddle:
		return "right", "center"
	default:
		return "left", "top" // Default alignment for unspecified cases
	}
}
