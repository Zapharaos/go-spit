# Styling, Borders & Merging

This page covers the visual customization options for **XLSX** export: styles, borders, cell
merging, and per-row/per-cell overrides.

!!! note
    Styling, borders and merging apply to XLSX output only. CSV is a plain-text format and
    ignores these options (it still supports hierarchical headers).

## Styles

A `Style` describes text and background appearance:

```go
type Style struct {
	Bold            bool      // Whether text should be bold
	Italic          bool      // Whether text should be italic
	Underline       string    // Underline style (format-specific values)
	TextColor       string    // Text color (hex, e.g. "#RRGGBB")
	BackgroundColor string    // Background color (hex, e.g. "#RRGGBB")
	FontSize        float64   // Font size in points
	FontFamily      string    // Font family name (e.g. "Arial")
	Alignment       Alignment // Text alignment
}
```

Apply a style to a column with `WithStyle`:

```go
spit.NewColumn("name", "Employee Name").
	WithStyle(&spit.Style{
		Bold:            true,
		BackgroundColor: "#FFE6E6",
		Alignment:       spit.AlignmentCenterMiddle,
	})
```

### Alignment

`Alignment` combines horizontal and vertical positioning:

| Constant                | Horizontal | Vertical |
|-------------------------|------------|----------|
| `AlignmentNone`         | default    | default  |
| `AlignmentLeft`         | left       | top      |
| `AlignmentCenter`       | center     | top      |
| `AlignmentRight`        | right      | top      |
| `AlignmentTop`          | left       | top      |
| `AlignmentMiddle`       | left       | center   |
| `AlignmentBottom`       | left       | bottom   |
| `AlignmentCenterMiddle` | center     | center   |
| `AlignmentLeftMiddle`   | left       | center   |
| `AlignmentRightMiddle`  | right      | center   |

## Borders

Borders are described per edge. A `Border` has a single `BorderStyle`, and `Borders` groups the
four edges plus optional inner borders.

### Border styles

| Constant            | Appearance              |
|---------------------|-------------------------|
| `BorderStyleNone`   | No border               |
| `BorderStyleThin`   | Thin solid line         |
| `BorderStyleMedium` | Medium solid line       |
| `BorderStyleDashed` | Dashed line             |
| `BorderStyleDotted` | Dotted line             |
| `BorderStyleThick`  | Thick solid line        |
| `BorderStyleDouble` | Double line             |

### Constructing borders

```go
// Same style on all four edges.
spit.NewBordersBoundaries(spit.BorderStyleThin)

// Different style per edge (left, right, top, bottom).
spit.NewBorders(
	spit.BorderStyleThin,   // left
	spit.BorderStyleThin,   // right
	spit.BorderStyleThick,  // top
	spit.BorderStyleThick,  // bottom
)
```

`Borders` also exposes chainable setters so you can build configurations incrementally:

| Method                     | Effect                                                |
|----------------------------|--------------------------------------------------------|
| `SetBoundaries(style)`     | All four outer edges.                                  |
| `SetVertical(style)`       | Left and right edges.                                  |
| `SetHorizontal(style)`     | Top and bottom edges.                                  |
| `SetLeft/Right/Top/Bottom` | A single edge.                                         |
| `SetInner(style)`          | Inner borders (used when a column has no sub-columns). |
| `HasBorders()`             | Reports whether any edge has a visible border.         |

```go
borders := spit.NewBordersBoundaries(spit.BorderStyleThin).
	SetInner(spit.BorderStyleDouble)

spit.NewColumn("department", "Department").WithBorders(borders)
```

## Merging

Cell merging combines adjacent cells that satisfy a condition. Conditions are defined by
`MergeCondition`:

| Constant                  | Merges when…                              |
|---------------------------|-------------------------------------------|
| `MergeConditionIdentical` | adjacent values are identical and non-empty |
| `MergeConditionEmpty`     | adjacent values are both empty or nil     |

`MergeRules` specifies which conditions apply vertically (between rows) and horizontally (between
columns). Build them with `NewMergeRules(vertical, horizontal)`:

```go
// Merge a column vertically when consecutive values are identical.
rules := spit.NewMergeRules(
	spit.MergeConditions{spit.MergeConditionIdentical}, // vertical
	nil, // horizontal
)

spit.NewColumn("department", "Department").WithMerge(rules)
```

Pass `nil` for a direction to disable merging in that direction.

## Header options

By default, headers use a bold, grey, centered style with thin borders. Override them with
`HeaderOptions`:

```go
table := spit.NewTable(data, columns, true).
	WithHeaderOptions(
		spit.NewHeaderOptions().
			WithStyle(&spit.Style{
				Bold:            true,
				BackgroundColor: "#D9E1F2",
				Alignment:       spit.AlignmentCenterMiddle,
			}).
			WithBorders(spit.NewBordersBoundaries(spit.BorderStyleMedium)),
	)
```

## Row options

`RowOptions` override styling, borders and merging for an entire data row. They are stored in a
`RowOptionsMap` keyed by the **0-based data row index**:

```go
rowOptions := spit.RowOptionsMap{
	1: *spit.NewRowOptions(1).
		WithStyle(&spit.Style{
			BackgroundColor: "#FFE6E6",
			Alignment:       spit.AlignmentCenterMiddle,
		}).
		WithBorder(spit.NewBordersBoundaries(spit.BorderStyleThick)),
}

table := spit.NewTable(data, columns, true).WithRowOptions(rowOptions)
```

Available builders: `WithStyle`, `WithBorder`, `WithMerge` and `WithMergeable`.

## Cell options

`CellOptions` provide the finest level of control, overriding both column and row settings for a
single cell. They are stored in a `CellOptionsMap`, a nested map keyed first by **column index**
and then by **row index**:

```go
cellOptions := spit.CellOptionsMap{
	4: { // column index 4
		5: *spit.NewCellOptions(5, 4). // row 5, column 4
			WithStyle(&spit.Style{
				BackgroundColor: "#FFFF99",
				Alignment:       spit.AlignmentCenterMiddle,
			}).
			WithBorder(spit.NewBordersBoundaries(spit.BorderStyleThick)).
			WithMergeable(false),
	},
}

table := spit.NewTable(data, columns, true).WithCellOptions(cellOptions)
```

Available builders: `WithStyle`, `WithBorder` and `WithMergeable`.

## Precedence

When several options apply to the same cell, the most specific configuration wins:

```text
Cell options  >  Row options  >  Column options  >  Defaults
```

For a complete, runnable demonstration combining all of these features, see the
[`examples/xlsx`](https://github.com/Zapharaos/go-spit/tree/main/examples/xlsx) program.
