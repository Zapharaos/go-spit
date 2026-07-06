# Tables, Data & Columns

Every export in go-spit revolves around a [`Table`](#tables): a combination of **data rows**,
**column definitions** and optional **styling/merging options**. This page explains each piece of
the data model.

## Data

A single row is a `Data` value — a map from a column name to its value:

```go
type Data map[string]interface{}
```

A collection of rows is a `DataSlice`:

```go
type DataSlice []Data
```

Example:

```go
data := spit.DataSlice{
	{"name": "John Doe", "age": 30, "salary": 75000.50},
	{"name": "Jane Smith", "age": 28, "salary": 82000},
}
```

Values can be any Go type. Numbers, booleans and `time.Time` values are handled natively;
other types are rendered using their default string representation.

### Nested data and lookups

`Data` supports nested maps. Use `Lookup` to read a (possibly nested) key:

```go
row := spit.Data{
	"address": spit.Data{"city": "Paris"},
}

value, err, found := row.Lookup("address", "city") // "Paris", nil, true
```

`Lookup` returns `(value, err, found)`. A missing key returns `found == false` with a `nil`
error; a malformed nested structure returns an error.

## Columns

A `Column` maps a data key to a header label and carries optional formatting and styling:

```go
type Column struct {
	Name    string      // Field name in the data source (for leaf columns)
	Label   string      // Display label for headers
	Format  string      // Format specification for value processing (e.g., date format)
	Width   float64     // Optional column width in character units (0 = use default)
	Merge   *MergeRules // Optional merge configuration for this column
	Borders *Borders    // Borders configuration
	Style   *Style      // Optional content style
	Columns Columns     // Sub-columns for hierarchical structures
}
```

Create columns with `NewColumn(name, label)` and configure them through a fluent API:

```go
columns := spit.Columns{
	spit.NewColumn("name", "Full Name"),
	spit.NewColumn("created_at", "Created At").WithFormat(time.RFC1123Z),
}
```

The available builder methods are:

| Method                       | Purpose                                                       |
|------------------------------|---------------------------------------------------------------|
| `WithFormat(format)`         | Set a value format (e.g. a date layout or an XLSX format key). |
| `WithWidth(width)`           | Set the column width in character units (0 = use default 15). |
| `WithStyle(style)`           | Apply a [`Style`](styling.md#styles) to the column's cells.   |
| `WithBorders(borders)`       | Apply [`Borders`](styling.md#borders) to the column's cells.  |
| `WithMerge(rules)`           | Apply [`MergeRules`](styling.md#merging) to the column.       |
| `WithSubColumns(subColumns)` | Replace the sub-columns (hierarchical headers).               |
| `AddSubColumn(subColumn)`    | Append a single sub-column.                                   |
| `RemoveSubColumn(name)`      | Remove a sub-column by name.                                  |

!!! note "Formatting"
    `Format` controls how a value is rendered. For dates, use a Go time layout such as
    `"2006-01-02"`. For XLSX-specific behaviors (formulas, hyperlinks, raw values, number/boolean
    coercion), use the
    [Excelize format constants](xlsx-export.md#cell-content-formats).

### Hierarchical (grouped) columns

Columns can be nested to create grouped, multi-level headers. A column with sub-columns acts as a
parent header that spans all of its leaf columns; only **leaf** columns map to data keys.

```go
columns := spit.Columns{
	spit.NewColumn("name", "Employee Name"),
	spit.NewColumn("", "Personal Information").
		WithSubColumns(spit.Columns{
			spit.NewColumn("age", "Age"),
			spit.NewColumn("email", "Email Address"),
		}),
}
```

This renders headers like:

```text
| Employee Name | Personal Information |
|               |   Age   |   Email    |
```

Both CSV and XLSX exports honor this hierarchy. In CSV, each header level becomes its own row; in
XLSX, parent headers are written above their children.

Useful helpers on `Columns` and `Column`:

| Helper                          | Description                                                  |
|---------------------------------|--------------------------------------------------------------|
| `Columns.GetFlattenedColumns()` | All leaf columns in order (used to read data values).        |
| `Columns.GetMaxDepth()`         | The number of header levels in the hierarchy.                |
| `Columns.GetTotalColumnCount()` | The total number of leaf columns.                            |
| `Column.HasSubColumns()`        | Whether a column has nested sub-columns.                     |
| `Column.CountSubColumns()`      | The number of leaf columns a column represents.              |

## Tables

A `Table` ties everything together:

```go
type Table struct {
	Data           DataSlice      // The actual data rows to be exported
	Columns        Columns        // Column definitions including hierarchy and formatting
	RowOptionsMap  RowOptionsMap  // Row-specific options (styling, merging, borders)
	CellOptionsMap CellOptionsMap // Cell-specific options for fine-grained control
	HeaderOptions  *HeaderOptions // Optional header configuration (style and borders)
	Preamble       PreambleRows   // Optional free-form rows written above the header/data area
	WriteHeader    bool           // Whether to generate headers from column definitions
	Limit          int64          // Maximum number of data rows to export (0 = no limit)
	ListSeparator  string         // Separator used when rendering slice/array values as strings
}
```

Build a table with `NewTable(data, columns, writeHeader)`. The third argument controls whether
headers are generated from the column definitions:

```go
table := spit.NewTable(data, columns, true)
```

Tables expose a fluent API to attach optional configuration:

| Method                          | Purpose                                                        |
|---------------------------------|----------------------------------------------------------------|
| `WithRowOptions(rowOptions)`    | Per-row styling, borders and merge overrides.                  |
| `WithCellOptions(cellOptions)`  | Per-cell styling, borders and merge overrides.                 |
| `WithHeaderOptions(options)`    | Override the default header style and borders.                 |
| `WithPreamble(preamble)`        | Prepend free-form rows above the header/data area.             |

```go
table := spit.NewTable(data, columns, true).
	WithRowOptions(rowOptions).
	WithCellOptions(cellOptions)
```

Row and cell options are most relevant for XLSX export; see
[Styling, Borders & Merging](styling.md) for details.

### Preamble rows

Preamble rows are free-form rows written **above** the table header. They are useful for adding
report titles, generation timestamps, or any other metadata that should appear at the top of the
sheet.

A `PreambleRow` carries an ordered list of cell values and an optional `Style`:

```go
type PreambleRow struct {
    Values []interface{} // Cell values (one entry per column position)
    Style  *Style        // Optional style applied to every non-empty cell
}
```

Build preamble rows with `NewPreambleRow` and attach them to a table with `WithPreamble`:

```go
table := spit.NewTable(data, columns, true).
    WithPreamble(spit.PreambleRows{
        spit.NewPreambleRow("Budget Report", "2024").
            WithStyle(&spit.Style{Bold: true}),
        spit.NewPreambleRow(), // empty spacer row
    })
```

Values are written left-to-right starting at column 1. Columns beyond the end of `Values` are
left empty. The header (and data rows) are automatically shifted down to accommodate the preamble.

!!! note
    Preamble rows apply to XLSX output only. CSV export ignores them.

### Rendering list values

When a cell value is a slice (`[]interface{}`), set `Table.ListSeparator` to control how the
elements are joined into a single string:

```go
table := spit.NewTable(data, columns, true)
table.ListSeparator = ", "
```

### Images

A cell value can be an `Image`. Each backend renders it in a format-appropriate way, so the same
data works across formats:

```go
data := spit.DataSlice{
	// Reference an external URL (HTML) or a local file path (XLSX).
	{"logo": spit.NewImageURL("https://acme.com/logo.png").WithAltText("Acme")},
	// Or embed binary content directly.
	{"logo": spit.NewImageBytes(pngBytes, "image/png").WithAltText("Globex").WithSize(48, 48)},
}
```

| Field     | Purpose                                                                   |
|-----------|---------------------------------------------------------------------------|
| `URL`     | Remote URL (HTML) or local file path (XLSX).                             |
| `Bytes`   | Embedded binary content.                                                  |
| `MIME`    | MIME type for embedded content (e.g. `"image/png"`); required with `Bytes`. |
| `AltText` | Alternative text; also the CSV/text fallback when no URL is set.         |
| `Width`, `Height` | Optional size hints in pixels (HTML output only).                |

Behavior per format:

| Format   | Rendering                                                                       |
|----------|---------------------------------------------------------------------------------|
| **HTML** | An `<img>` element — data URI for embedded bytes, otherwise `src=URL`.           |
| **XLSX** | A cell-anchored picture (auto-fit). `URL` must be a **local file path**; use `Bytes` for remote images. |
| **CSV**  | Text fallback: the `URL` (or `AltText` when no URL is set).                      |

!!! note
    For XLSX, remote URLs are **not** downloaded — Excelize inserts pictures from a local file path
    or from bytes. Provide `Bytes` (with `MIME`) to embed a remote or in-memory image in a workbook.
