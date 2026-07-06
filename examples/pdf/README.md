# HTML → PDF pipeline

This example shows the recommended way to produce **good-looking PDFs** with go-spit.

go-spit is a pure-Go library and deliberately does **not** embed a PDF engine: producing
print-quality PDFs requires a real layout/typesetting engine, and no pure-Go engine reaches that
quality today. Instead, the pipeline is:

1. **go-spit generates print-ready HTML/CSS** (pure Go, no extra dependency) — `@page` rules,
   controlled page breaks, and a repeated table header.
2. **A rendering engine turns that HTML into a PDF.** The quality of the PDF is the quality of the
   HTML/CSS — exactly how most modern services generate invoices and reports.

This keeps `go get github.com/Zapharaos/go-spit` lightweight: the heavy engine is the consumer's
choice, never imposed by the library.

## Running this example

```sh
go run .
```

It writes `invoice.html`, then converts it to `invoice.pdf` using a **Chromium-based browser already
installed on your machine** (Microsoft Edge, Google Chrome or Chromium) in headless mode, invoked via
`os/exec`. No Go dependency is added. If no browser is found, it tells you to print the HTML
manually or plug in another engine.

The browser call is essentially:

```sh
msedge --headless=new --disable-gpu --no-pdf-header-footer \
  --print-to-pdf=invoice.pdf "file:///…/invoice.html"
```

## What makes the PDF look good

The `printCSS` in [main.go](main.go) is where the print quality comes from:

- `@page { size: A4; margin: 18mm; }` — page geometry.
- `page-break-inside: avoid` on rows/sections — no ugly splits across pages.
- `thead { display: table-header-group; }` — the table header repeats on every page the table spans.

## Other engines (pick per your environment)

The same generated HTML can be fed to any HTML/CSS → PDF engine:

| Engine | Type | Notes |
|--------|------|-------|
| **Headless Chromium** (Edge/Chrome, or `chromedp`/`go-rod`) | Browser | Free, ubiquitous, great fidelity. Used here. |
| **Gotenberg** | Docker service | Wraps Chromium (+ LibreOffice); popular self-hosted microservice. Send the HTML over HTTP. |
| **PrinceXML** / **DocRaptor** | Dedicated print engine | Best-in-class print CSS (running headers, TOC, hyphenation). Commercial. |
| **WeasyPrint** | Print engine | Open source (Python), no browser needed. |
| **Antenna House**, **PDFreactor** | Print engines | High-end commercial. |

For a turnkey Go integration, a thin `chromedp` adapter could live in a **separate module**
(its own `go.mod`) so importing go-spit never pulls the dependency — only users who want it opt in.
