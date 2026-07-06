# HTML Export

go-spit renders tabular data as a styled HTML `<table>`, and can compose full documents with
headings, paragraphs, lists and sections around one or more tables. The HTML backend reuses the
exact same styling and merging model as the [XLSX export](xlsx-export.md): per-cell, per-column,
per-row and header `Style` values become inline CSS, and cell merges become `rowspan`/`colspan`.
Document-level presentation — title, description, page font/background, custom CSS — is configured
through `HTMLOptions`.

## The export functions

```go
// Render a single table.
func ExportHTML(t *Table, opts HTMLOptions, params FileWriteParams) (*FileWriteResult, error)

// Render a composed document (headings, paragraphs, lists, sections, tables).
func ExportHTMLDocument(doc *HTMLDocument, params FileWriteParams) (*FileWriteResult, error)
```

The `.html` extension is added automatically when `params.Extension` is empty. See
[File Options](file-options.md) for the available `params`. For multi-block documents, jump to
[Composing a full document](#composing-a-full-document).

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

## Composing a full document

`ExportHTML` renders a single table. To build a complete document — headings, paragraphs, lists,
sections and one or more tables — use `HTMLDocument` and `ExportHTMLDocument`. The document shares
the same [`HTMLOptions`](#document-options) and reuses the full table styling/merging.

```go
salesTable := spit.NewTable(salesData, salesColumns, true)

doc := spit.NewHTMLDocument(spit.HTMLOptions{
	Title:       "Annual Report",
	Description: "Fiscal year 2026",
	BodyStyle:   &spit.Style{FontFamily: "Segoe UI"},
}).
	Heading(2, "Summary").
	Paragraph("Revenue grew across every region.").
	UnorderedList("Two new markets", "Team doubled").
	Section(2, "Sales",
		spit.Paragraph("Quarterly breakdown:"),
		spit.TableBlock(salesTable),
		spit.OrderedList("Q1 steady", "Q2 up 25%"),
	)

result, err := spit.ExportHTMLDocument(doc, spit.FileWriteParams{Filename: "report"})
if err != nil {
	log.Fatal(err)
}
defer result.RemoveFile()
```

### Blocks

A document is an ordered list of blocks. Each constructor returns an `HTMLBlock`; the fluent
`HTMLDocument` methods (`Heading`, `Paragraph`, …) are shortcuts for `Add(<block>)`.

| Block                          | Renders                                                             |
|--------------------------------|---------------------------------------------------------------------|
| `Heading(level, text)`         | `<h1>`–`<h6>` (level clamped to 1–6) — titles and subtitles.        |
| `Paragraph(text)`              | `<p>`.                                                               |
| `UnorderedList(items...)`      | `<ul>` with one `<li>` per item.                                     |
| `OrderedList(items...)`        | `<ol>` with one `<li>` per item.                                     |
| `DefinitionList(items...)`     | `<dl>` with `<dt>`/`<dd>` pairs (use `Def(term, desc)`).            |
| `Blockquote(text)`             | `<blockquote>`.                                                      |
| `CodeBlock(code)`              | `<pre><code>` — preformatted, escaped.                              |
| `HorizontalRule()`             | `<hr>` — a thematic break.                                           |
| `ImageBlock(img)`              | A standalone `<img>` (reuses the [`Image`](tables-and-columns.md#images) type). |
| `TableBlock(table)`            | A `<table>` with the table's full styling, borders and merging.     |
| `Section(level, title, ...)`   | A semantic `<section>` with a heading, wrapping nested blocks.       |
| `RawHTML(markup)`              | The given markup **verbatim** (not escaped) — an escape hatch.       |

All text (`Heading`, `Paragraph`, list items, `Section` titles, etc.) is HTML-escaped. Only
`RawHTML` emits unescaped content, so pass it trusted markup only. Document `Title`/`Description`
from `HTMLOptions` still render at the top, before the blocks. `FragmentOnly` works the same way as
for single-table export.

**Per-block styling.** `Heading`, `Paragraph`, `Blockquote`, lists and `Section` accept a
`WithStyle(*Style)` for an inline style. `TableBlock` accepts `WithCaption(string)` (an accessible
`<caption>`) and `WithStyle(*Style)` (overrides the document `TableStyle` for that table):

```go
spit.Heading(2, "Overview").WithStyle(&spit.Style{TextColor: "#0969DA"})
spit.TableBlock(salesTable).WithCaption("Q3 sales")
```

**Nested lists.** Build sub-items with `Item(text, children...)` and attach them with the list's
`Add` method:

```go
spit.UnorderedList("Fruits").
	Add(spit.Item("Citrus", spit.Item("Orange"), spit.Item("Lemon")))
```

### Table of contents and anchors

Set `HTMLOptions.TableOfContents` to render a linked `<nav class="toc">` from the document's
headings and section titles. Enabling it also adds a slugified `id` to every heading (deduplicated
across the document), so headings become deep-linkable. When it is off (the default), no ids are
emitted.

### Theme

`HTMLOptions.Theme` injects a built-in stylesheet for a polished look with zero styling effort:

| Value              | Effect                                                                        |
|--------------------|-------------------------------------------------------------------------------|
| `HTMLThemeNone`    | No stylesheet (default). Only explicit styles are applied.                    |
| `HTMLThemeDefault` | A clean, readable stylesheet: modern font stack, spacing, styled headings, zebra-striped tables, responsive max width. |

`CustomCSS` is injected after the theme, so it can override any theme rule. When a theme is active,
cell padding is left to the stylesheet (the inline default is dropped).

### Table structure

Rendered tables use semantic `<thead>`/`<tbody>` grouping, map each column's
[`Width`](tables-and-columns.md) to a `<colgroup>`/`<col>` (`ch` units), and right-align numeric
data cells automatically (unless the cell has an explicit alignment).

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
| `Theme`        | `HTMLTheme` | Built-in stylesheet for a polished default look (see [Theme](#theme)). Default: `HTMLThemeNone`.      |
| `TableOfContents` | `bool` | Documents only: render a linked table of contents from the headings (see [above](#table-of-contents-and-anchors)). |

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
