# API Reference

The complete, always-up-to-date Go API reference is generated from the source code and hosted on
**pkg.go.dev**:

[:material-open-in-new: Browse the full API on pkg.go.dev](https://pkg.go.dev/github.com/Zapharaos/go-spit){ .md-button .md-button--primary }

## Package overview

The package is imported as `spit`:

```go
import "github.com/Zapharaos/go-spit"
```

### Exporters

| Symbol                       | Description                                        |
|------------------------------|----------------------------------------------------|
| `ExportCSV`                  | Export a table to a CSV file.                      |
| `ExportXLSX`                 | Export a single sheet to an XLSX file.             |
| `ExportXLSXSheets`           | Export multiple sheets to one XLSX workbook.       |
| `ExportHTML`                 | Export a table to a styled HTML document.          |
| `ExportHTMLDocument`         | Export a composed HTML document (headings, paragraphs, lists, sections, tables). |

Google Sheets export lives in the optional [`gsheets`](../user-guide/google-sheets.md) module
(`gsheets.ExportGoogleSheets`), kept separate so the core package stays dependency-light.

### Data model

| Symbol                            | Description                                  |
|-----------------------------------|----------------------------------------------|
| `Table`, `NewTable`               | The table to export.                         |
| `Data`, `DataSlice`               | Row data structures.                         |
| `Column`, `Columns`, `NewColumn`  | Column definitions and hierarchies.          |
| `HeaderOptions`, `NewHeaderOptions` | Header style/border overrides.             |
| `PreambleRow`, `PreambleRows`, `NewPreambleRow` | Free-form rows above the header. |
| `RowOptions`, `RowOptionsMap`     | Per-row overrides.                           |
| `CellOptions`, `CellOptionsMap`   | Per-cell overrides.                          |
| `HTMLOptions`                     | Document-level options for HTML export (title, description, page styling). |
| `HTMLDocument`, `NewHTMLDocument` | Composed HTML document (a sequence of blocks).      |
| `HTMLTheme`                       | Built-in HTML stylesheet selector (`HTMLThemeNone`, `HTMLThemeDefault`). |
| `HTMLBlock`, `Heading`, `Paragraph`, `UnorderedList`, `OrderedList`, `DefinitionList`, `Blockquote`, `CodeBlock`, `HorizontalRule`, `ImageBlock`, `TableBlock`, `Section`, `RawHTML` | HTML document block constructors. |
| `Item`, `ListItem`, `Def`, `DefinitionItem` | Helpers for nested list items and definition entries. |
| `Image`, `NewImageURL`, `NewImageBytes` | Image cell values (HTML/XLSX render them; CSV falls back to text). |

### Styling

| Symbol                                   | Description                          |
|------------------------------------------|--------------------------------------|
| `Style`, `Alignment`                     | Text and background styling.         |
| `Border`, `Borders`, `BorderStyle`       | Border configuration.                |
| `MergeRules`, `MergeConditions`, `MergeCondition` | Cell merging rules.         |

### Spreadsheets

| Symbol                                       | Description                          |
|----------------------------------------------|--------------------------------------|
| `Spreadsheet`                                | Backend-agnostic spreadsheet interface. |
| `SpreadsheetExcelize`, `NewSpreadsheetExcelize` | Excelize-backed implementation.   |
| `ExcelizeFormatDefault/Formula/Hyperlink`    | XLSX cell content formats.           |

### Files

| Symbol                                  | Description                            |
|-----------------------------------------|----------------------------------------|
| `FileWriteParams`, `FileWriteResult`    | File writing inputs and results.       |
| `SanitizeFilename`                      | Make a string safe to use as a filename. |

### Utilities & logging

| Symbol                                          | Description                       |
|-------------------------------------------------|-----------------------------------|
| `FormatValue`, `ConvertSliceToString`, `ParseDate` | Value formatting helpers.      |
| `Format`                                        | Export format identifier.         |
| `Logger`, `Field`, `StdLogger`                  | Logging interface and helpers.    |
| `SetLogger`, `SetLogLevel`, `GetLogLevel`, `HasLogLevel`, `DisableLogger`, `ResetLogger` | Logger configuration. |
