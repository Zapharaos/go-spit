// pdf demonstrates the recommended "quality PDF" pipeline for go-spit:
//
//  1. go-spit generates a print-ready HTML document (pure Go, no extra dependency).
//  2. A real rendering engine turns that HTML into a PDF.
//
// This example uses a Chromium-based browser already installed on the machine
// (Microsoft Edge, Google Chrome or Chromium) in headless mode, invoked through
// os/exec — so it adds NO Go dependency to go-spit. The browser is the industry-
// standard way to get good-looking PDFs from HTML/CSS. See README.md for other
// engines (Gotenberg, PrinceXML, WeasyPrint, chromedp).
package main

import (
	"context"
	"fmt"
	"log"
	"os"
	"os/exec"
	"path/filepath"
	"time"

	"github.com/Zapharaos/go-spit"
)

// printCSS turns the on-screen document into a print-optimized one. These rules are
// what actually make the resulting PDF look professional; the rendering engine (the
// browser) does the hard layout work.
const printCSS = `
@page { size: A4; margin: 18mm; }
@media print {
  body { max-width: none; margin: 0; }
  /* Keep headings with their content and avoid splitting rows/sections across pages. */
  h1, h2, h3 { page-break-after: avoid; }
  tr, section, blockquote { page-break-inside: avoid; }
  /* Repeat the table header on every page the table spans. */
  thead { display: table-header-group; }
}
`

// buildReport composes a print-ready document with go-spit.
func buildReport() *spit.HTMLDocument {
	// Invoice-style line items; numeric columns are right-aligned automatically.
	items := spit.DataSlice{
		{"desc": "Design services", "qty": 12, "unit": 85.0, "total": 1020.0},
		{"desc": "Frontend development", "qty": 40, "unit": 95.0, "total": 3800.0},
		{"desc": "Backend development", "qty": 32, "unit": 105.0, "total": 3360.0},
		{"desc": "QA & testing", "qty": 16, "unit": 75.0, "total": 1200.0},
	}
	columns := spit.Columns{
		spit.NewColumn("desc", "Description").WithWidth(40),
		spit.NewColumn("qty", "Qty").WithWidth(8),
		spit.NewColumn("unit", "Unit (€)").WithWidth(12),
		spit.NewColumn("total", "Total (€)").WithWidth(12),
	}
	lineItems := spit.NewTable(items, columns, true)

	return spit.NewHTMLDocument(spit.HTMLOptions{
		Title:           "Invoice #2026-0042",
		Description:     "Acme Corp — issued 6 July 2026",
		Theme:           spit.HTMLThemeDefault,
		TableOfContents: true,
		CustomCSS:       printCSS,
	}).
		Section(2, "Billing details",
			spit.DefinitionList(
				spit.Def("Bill to", "Globex Industries"),
				spit.Def("Invoice date", "2026-07-06"),
				spit.Def("Due date", "2026-08-05"),
				spit.Def("Payment terms", "Net 30"),
			),
		).
		Section(2, "Line items",
			spit.Paragraph("Work performed during June 2026:"),
			spit.TableBlock(lineItems).WithCaption("Billable work"),
		).
		Section(2, "Notes",
			spit.Blockquote("Thank you for your business. Please reference the invoice number with your payment."),
			spit.UnorderedList(
				"Bank transfer preferred",
				"Late payments subject to 1.5% monthly interest",
			),
		)
}

func main() {
	// Step 1 — generate the print-ready HTML with go-spit.
	doc := buildReport()
	result, err := spit.ExportHTMLDocument(doc, spit.FileWriteParams{
		Filename:      "invoice",
		OverwriteFile: true,
	})
	if err != nil {
		log.Fatalf("failed to generate HTML: %v", err)
	}
	htmlPath, err := filepath.Abs(result.Filepath)
	if err != nil {
		log.Fatal(err)
	}
	log.Printf("HTML written to %s", htmlPath)

	// Step 2 — convert HTML to PDF with a headless Chromium-based browser.
	browser := findBrowser()
	if browser == "" {
		log.Printf("No Chromium-based browser (Edge/Chrome/Chromium) found on this machine.")
		log.Printf("Open %s in a browser and use \"Print > Save as PDF\", or plug in another engine (see README.md).", htmlPath)
		return
	}

	pdfPath := filepath.Join(filepath.Dir(htmlPath), "invoice.pdf")
	if err := htmlToPDF(browser, htmlPath, pdfPath); err != nil {
		log.Fatalf("PDF conversion failed using %s: %v", browser, err)
	}
	log.Printf("PDF written to %s (rendered by %s)", pdfPath, filepath.Base(browser))
}

// htmlToPDF renders a local HTML file to a PDF using a headless Chromium-based browser.
func htmlToPDF(browser, htmlPath, pdfPath string) error {
	// Chromium accepts a file:// URL as the page to print.
	fileURL := "file:///" + filepath.ToSlash(htmlPath)

	// Use a throwaway profile dir so we never touch the user's real browser profile.
	profileDir, err := os.MkdirTemp("", "gospit-pdf-")
	if err != nil {
		return err
	}
	defer os.RemoveAll(profileDir)

	ctx, cancel := context.WithTimeout(context.Background(), 60*time.Second)
	defer cancel()

	args := []string{
		"--headless=new",
		"--disable-gpu",
		"--no-pdf-header-footer", // drop the default date/URL header & footer
		"--user-data-dir=" + profileDir,
		"--print-to-pdf=" + pdfPath,
		fileURL,
	}
	cmd := exec.CommandContext(ctx, browser, args...)
	out, err := cmd.CombinedOutput()
	if err != nil {
		return fmt.Errorf("%w: %s", err, out)
	}
	if _, statErr := os.Stat(pdfPath); statErr != nil {
		return fmt.Errorf("browser exited without producing a PDF: %s", out)
	}
	return nil
}

// findBrowser locates a Chromium-based browser executable, checking PATH and common
// install locations across platforms. Returns "" when none is found.
func findBrowser() string {
	// First, anything already on PATH.
	for _, name := range []string{
		"msedge", "chrome", "google-chrome", "chromium", "chromium-browser",
		"msedge.exe", "chrome.exe",
	} {
		if p, err := exec.LookPath(name); err == nil {
			return p
		}
	}

	// Then well-known absolute install paths.
	candidates := []string{
		// Windows — Microsoft Edge is present on every Windows 10/11 machine.
		`C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe`,
		`C:\Program Files\Microsoft\Edge\Application\msedge.exe`,
		`C:\Program Files\Google\Chrome\Application\chrome.exe`,
		`C:\Program Files (x86)\Google\Chrome\Application\chrome.exe`,
		// macOS
		"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
		"/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
		"/Applications/Chromium.app/Contents/MacOS/Chromium",
		// Linux
		"/usr/bin/google-chrome",
		"/usr/bin/chromium",
		"/usr/bin/chromium-browser",
		"/usr/bin/microsoft-edge",
	}
	for _, c := range candidates {
		if _, err := os.Stat(c); err == nil {
			return c
		}
	}
	return ""
}
