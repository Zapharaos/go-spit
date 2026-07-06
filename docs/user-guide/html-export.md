# HTML Export

go-spit renders tabular data as a styled HTML `<table>`. The HTML backend reuses the exact same
styling and merging model as the [XLSX export](xlsx-export.md): per-cell, per-column, per-row and
header `Style` values become inline CSS, and cell merges become `rowspan`/`colspan`. Document-level
presentation — title, description, page font/background, custom CSS — is configured through
`HTMLOptions`.

## The export function

```go
func ExportHTML(t *Table, opts HTMLOptions, params FileWriteParams) (*FileWriteResult, error)
```

The `.html` extension is added automatically when `params.Extension` is empty. See
[File Options](file-options.md) for the available `params`.

## Basic example

```go
package main

import (
	"log"

	"github.com/Zapharaos/go-spit"
)

func main() {
	data := spit.DataSlice{
		{"name": "John Doe", "age": 30, "department": "Engineering"},
		{"name": "Jane Smith", "age": 28, "department": "Engineering"},
	}

	columns := spit.Columns{
		spit.NewColumn("name", "Employee Name").
			WithStyle(&spit.Style{Bold: true, TextColor: "#1155CC"}),
		spit.NewColumn("age", "Age"),
		spit.NewColumn("department", "Department").
			WithMerge(spit.NewMergeRules(spit.MergeConditions{spit.MergeConditionIdentical}, nil)),
	}

	table := spit.NewTable(data, columns, true)

	opts := spit.HTMLOptions{
		Title:       "Employee Report",
		Description: "Generated with go-spit.",
		BodyStyle:   &spit.Style{FontFamily: "Segoe UI", BackgroundColor: "#FAFAFA"},
	}

	result, err := spit.ExportHTML(table, opts, spit.FileWriteParams{
		Filename: "employee_report",
	})
	if err != nil {
		log.Fatal(err)
	}
	defer result.RemoveFile()

	log.Printf("created %s", result.Filepath)
}
```

## Document options

`HTMLOptions` carries only the presentation options that have no equivalent in the tabular model.
Everything about cell content and cell styling comes from the `Table` itself.

| Field          | Type     | Description                                                                                             |
|----------------|----------|---------------------------------------------------------------------------------------------------------|
| `Title`        | `string` | Rendered as an `<h1>` above the table, and as the `<title>` of full documents.                          |
| `Description`  | `string` | Rendered as a `<p>` below the title.                                                                     |
| `BodyStyle`    | `*Style` | Default style applied to the page `<body>` (font family, background, text color, …).                    |
| `TableStyle`   | `*Style` | Style applied to the `<table>` element itself. `border-collapse:collapse` is always set.                |
| `CustomCSS`    | `string` | Raw CSS injected into a `<style>` block for advanced customization.                                     |
| `Lang`         | `string` | `lang` attribute for the `<html>` element (default: `"en"`).                                            |
| `FragmentOnly` | `bool`   | When `true`, emit only the title/description/table markup without the document wrappers (see below).    |

## Full document vs. fragment

By default `ExportHTML` produces a standalone document (`<!DOCTYPE html>`, `<head>`, `<body>`) that
opens directly in a browser:

```html
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Employee Report</title>
</head>
<body style="background-color:#FAFAFA;font-family:'Segoe UI'">
<h1>Employee Report</h1>
<p>Generated with go-spit.</p>
<table style="border-collapse:collapse">…</table>
</body>
</html>
```

Set `FragmentOnly: true` to emit just the title, description and `<table>` — useful for embedding the
result into an existing page or an email body:

```go
opts := spit.HTMLOptions{Title: "Employee Report", FragmentOnly: true}
```

```html
<h1>Employee Report</h1>
<table style="border-collapse:collapse">…</table>
```

!!! tip "Email-friendly output"
    Cell styles, borders and background colors are emitted as **inline `style` attributes** on each
    `<td>`/`<th>`, not as CSS classes. This keeps fragments self-contained and portable across email
    clients that strip `<style>` blocks.

## Styling

HTML export reuses the shared [`Style`](styling.md) type, so cell, column, row and header styling
work exactly as they do for XLSX. Each `Style` field maps to CSS as follows:

| `Style` field     | CSS output                                                    |
|-------------------|---------------------------------------------------------------|
| `Bold`            | `font-weight:bold`                                            |
| `Italic`          | `font-style:italic`                                          |
| `Underline`       | `text-decoration:underline`                                  |
| `TextColor`       | `color:…` (bare 6-digit hex is prefixed with `#`)            |
| `BackgroundColor` | `background-color:…`                                         |
| `FontSize`        | `font-size:…pt`                                              |
| `FontFamily`      | `font-family:…` (quoted when it contains spaces)            |
| `Alignment`       | `text-align` + `vertical-align`                              |

!!! note
    `NumFmt` (the Excel number-format string) has no HTML equivalent and is ignored. Format values
    for display with a column `Format` instead — see [Tables, Data & Columns](tables-and-columns.md).

Borders defined via `Borders` / `WithBorders` map to `border-*` declarations. Because the table uses
`border-collapse:collapse`, adjacent borders render as single lines just like in a spreadsheet.

## Cell content formats

The `Format` field on a column behaves consistently with the other backends:

| Constant                  | Behavior in HTML                                                        |
|---------------------------|------------------------------------------------------------------------|
| `ExcelizeFormatHyperlink` | Renders the value as a clickable `<a href="…">` element.               |
| `ExcelizeFormatFormula`   | HTML has no formulas — the value is rendered as plain text.            |
| date layouts / custom     | The value is formatted and rendered as text.                          |

```go
columns := spit.Columns{
	spit.NewColumn("homepage", "Website").WithFormat(spit.ExcelizeFormatHyperlink),
}
```

## Images

Put an `Image` value into a cell to render an `<img>` element. Provide either a URL (or local
path) or embedded bytes; embedded content is emitted as a base64 data URI, keeping the document
self-contained:

```go
data := spit.DataSlice{
	{"name": "Acme", "logo": spit.NewImageURL("https://acme.com/logo.png").
		WithAltText("Acme").WithSize(48, 48)},
	{"name": "Globex", "logo": spit.NewImageBytes(pngBytes, "image/png").
		WithAltText("Globex")},
}

columns := spit.Columns{
	spit.NewColumn("name", "Company"),
	spit.NewColumn("logo", "Logo"),
}
```

`Width` and `Height` (via `WithSize`) are applied to HTML output as `width`/`height` attributes.
See [Images across formats](tables-and-columns.md#images) for how the same `Image` value behaves in
XLSX and CSV.

## Escaping & safety

All cell values, the title and the description are HTML-escaped, so untrusted data cannot inject
markup:

```go
// A value of `<script>alert(1)</script>` is rendered as
//   &lt;script&gt;alert(1)&lt;/script&gt;
```

!!! warning
    `CustomCSS` and hyperlink URLs are the two places where you control raw output. `CustomCSS` is
    injected verbatim into a `<style>` block, and hyperlink URLs are placed in an `href` attribute
    (attribute-escaped, but not validated). Only pass trusted values to these.

## Preamble, headers and merging

HTML export honors the same table features as XLSX:

- **Preamble rows** — free-form rows written above the header (`WithPreamble`).
- **Hierarchical headers** — nested columns produce multi-row headers with the appropriate
  `colspan`/`rowspan`.
- **Cell merging** — vertical and horizontal merging based on identical or empty values collapses
  into spanned `<td>`/`<th>` cells.

These are covered in detail in [Styling, Borders & Merging](styling.md) and
[Tables, Data & Columns](tables-and-columns.md).
