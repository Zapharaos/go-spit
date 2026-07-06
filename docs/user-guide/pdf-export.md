# Generating a PDF

go-spit is a pure-Go library and does **not** embed a PDF engine. Print-quality PDFs require a real
layout/typesetting engine, and no pure-Go engine reaches that quality today — so instead of shipping
a mediocre native renderer (or forcing a heavy dependency on every consumer), go-spit produces
**print-ready HTML/CSS** and lets you feed it to the engine of your choice.

This is how most modern services generate invoices and reports: the quality of the PDF is the
quality of the HTML/CSS.

```
go-spit (pure Go)  →  print-ready HTML  →  rendering engine  →  PDF
```

## Print-ready HTML

Build an [HTML document](html-export.md#composing-a-full-document) and add a small print stylesheet
via `CustomCSS`:

```go
const printCSS = `
@page { size: A4; margin: 18mm; }
@media print {
  body { max-width: none; margin: 0; }
  h1, h2, h3 { page-break-after: avoid; }
  tr, section { page-break-inside: avoid; }
  thead { display: table-header-group; } /* repeat table header on each page */
}
`

doc := spit.NewHTMLDocument(spit.HTMLOptions{
    Title:     "Invoice #2026-0042",
    Theme:     spit.HTMLThemeDefault,
    CustomCSS: printCSS,
}).
    Section(2, "Line items", spit.TableBlock(lineItems).WithCaption("Billable work"))

result, _ := spit.ExportHTMLDocument(doc, spit.FileWriteParams{Filename: "invoice"})
```

The key rules — `@page` geometry, `page-break-inside: avoid`, and `thead { display:
table-header-group }` (which repeats the table header on every page) — are what make the PDF look
professional. The rendering engine does the actual layout.

## Rendering to PDF

The same HTML works with any HTML/CSS → PDF engine; pick one per your environment:

| Engine | Type | Notes |
|--------|------|-------|
| Headless Chromium (Edge/Chrome, `chromedp`, `go-rod`) | Browser | Free, ubiquitous, great fidelity. |
| Gotenberg | Docker service | Wraps Chromium; common self-hosted microservice. |
| PrinceXML / DocRaptor | Print engine | Best print CSS (running headers, TOC). Commercial. |
| WeasyPrint | Print engine | Open source, no browser needed. |

The [`examples/pdf`](https://github.com/Zapharaos/go-spit/tree/main/examples/pdf) example runs the
full pipeline end to end: it generates the HTML, then converts it with a headless Chromium-based
browser already installed on the machine (invoked via `os/exec`, so it adds **no Go dependency**).

!!! note "Keeping go-spit lightweight"
    Because the engine lives outside the library, `go get github.com/Zapharaos/go-spit` stays pure
    Go with no external binary. For a turnkey Go integration, a thin `chromedp` adapter can live in a
    **separate module** (its own `go.mod`) so only users who want it pull the dependency.
