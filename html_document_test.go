package spit

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func sampleDocTable() *Table {
	return NewTable(
		DataSlice{{"q": "Q1", "rev": 120}, {"q": "Q2", "rev": 150}},
		Columns{NewColumn("q", "Quarter"), NewColumn("rev", "Revenue")},
		true,
	)
}

func renderDoc(t *testing.T, doc *HTMLDocument) string {
	t.Helper()
	out, err := doc.render()
	if err != nil {
		t.Fatalf("render failed: %v", err)
	}
	return out
}

func TestHTMLDocumentBlocks(t *testing.T) {
	doc := NewHTMLDocument(HTMLOptions{Title: "Report", Description: "2026"}).
		Heading(2, "Overview").
		Paragraph("Intro text").
		UnorderedList("a", "b").
		OrderedList("first", "second").
		Table(sampleDocTable())

	out := renderDoc(t, doc)

	for _, want := range []string{
		"<!DOCTYPE html>",
		"<h1>Report</h1>",
		"<p>2026</p>",
		"<h2>Overview</h2>",
		"<p>Intro text</p>",
		"<ul>\n<li>a</li>\n<li>b</li>\n</ul>",
		"<ol>\n<li>first</li>\n<li>second</li>\n</ol>",
		"<table style=\"border-collapse:collapse\">",
		">Quarter<",
		"</body>\n</html>",
	} {
		if !strings.Contains(out, want) {
			t.Errorf("output missing %q", want)
		}
	}
}

func TestHTMLDocumentSection(t *testing.T) {
	doc := NewHTMLDocument(HTMLOptions{}).
		Section(2, "Sales",
			Paragraph("breakdown"),
			TableBlock(sampleDocTable()),
		)
	out := renderDoc(t, doc)

	if !strings.Contains(out, "<section>\n<h2>Sales</h2>") {
		t.Errorf("expected section with heading, got:\n%s", out)
	}
	// The table must be nested inside the section.
	sectionStart := strings.Index(out, "<section>")
	sectionEnd := strings.Index(out, "</section>")
	tableIdx := strings.Index(out, "<table")
	if sectionStart < 0 || sectionEnd < 0 || tableIdx < sectionStart || tableIdx > sectionEnd {
		t.Error("table should be rendered inside the section")
	}
}

func TestHTMLDocumentEscaping(t *testing.T) {
	doc := NewHTMLDocument(HTMLOptions{}).
		Heading(3, "A <b> & C").
		Paragraph("x < y & z").
		UnorderedList("<script>")
	out := renderDoc(t, doc)

	for _, want := range []string{"<h3>A &lt;b&gt; &amp; C</h3>", "x &lt; y &amp; z", "<li>&lt;script&gt;</li>"} {
		if !strings.Contains(out, want) {
			t.Errorf("expected escaped %q", want)
		}
	}
	if strings.Contains(out, "<script>") {
		t.Error("unescaped script leaked")
	}
}

func TestHTMLDocumentRawHTML(t *testing.T) {
	doc := NewHTMLDocument(HTMLOptions{}).Add(RawHTML(`<div class="note">Hi</div>`))
	out := renderDoc(t, doc)
	if !strings.Contains(out, `<div class="note">Hi</div>`) {
		t.Error("raw HTML should be emitted verbatim")
	}
}

func TestHeadingLevelClamp(t *testing.T) {
	out := renderDoc(t, NewHTMLDocument(HTMLOptions{}).Heading(0, "lo").Heading(9, "hi"))
	if !strings.Contains(out, "<h1>lo</h1>") || !strings.Contains(out, "<h6>hi</h6>") {
		t.Errorf("heading levels not clamped to 1-6, got:\n%s", out)
	}
}

func TestHTMLDocumentFragmentOnly(t *testing.T) {
	doc := NewHTMLDocument(HTMLOptions{Title: "T", FragmentOnly: true}).
		Paragraph("body")
	out := renderDoc(t, doc)
	if strings.Contains(out, "<!DOCTYPE") || strings.Contains(out, "<body") {
		t.Errorf("fragment should omit document wrappers, got:\n%s", out)
	}
	if !strings.Contains(out, "<h1>T</h1>") || !strings.Contains(out, "<p>body</p>") {
		t.Error("fragment should still contain title and content")
	}
}

func TestExportHTMLDocumentWritesFile(t *testing.T) {
	dir := t.TempDir()
	doc := NewHTMLDocument(HTMLOptions{Title: "Doc"}).
		Heading(2, "Section").
		Paragraph("hello").
		Table(sampleDocTable())

	result, err := ExportHTMLDocument(doc, FileWriteParams{Filename: "doc", Filepath: dir, OverwriteFile: true})
	if err != nil {
		t.Fatalf("ExportHTMLDocument failed: %v", err)
	}
	if filepath.Ext(result.Filename) != ".html" {
		t.Errorf("expected .html extension, got %q", result.Filename)
	}
	content, err := os.ReadFile(result.Filepath)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(content), "<h1>Doc</h1>") || !strings.Contains(string(content), "<table") {
		t.Error("written document missing expected content")
	}
}

func TestExportHTMLDocumentNil(t *testing.T) {
	if _, err := ExportHTMLDocument(nil, FileWriteParams{Filename: "x"}); err == nil {
		t.Error("expected error for nil document")
	}
}

func TestHTMLDocumentContentBlocks(t *testing.T) {
	doc := NewHTMLDocument(HTMLOptions{}).
		Blockquote("quote").
		CodeBlock("go test ./...").
		HorizontalRule().
		DefinitionList(Def("Author", "Zap"), Def("Year", "2026")).
		Image(NewImageURL("https://x/y.png").WithAltText("pic"))
	out := renderDoc(t, doc)

	for _, want := range []string{
		"<blockquote>quote</blockquote>",
		"<pre><code>go test ./...</code></pre>",
		"<hr>",
		"<dl>\n<dt>Author</dt>\n<dd>Zap</dd>\n<dt>Year</dt>\n<dd>2026</dd>\n</dl>",
		`<img src="https://x/y.png" alt="pic">`,
	} {
		if !strings.Contains(out, want) {
			t.Errorf("output missing %q", want)
		}
	}
}

func TestHTMLNestedList(t *testing.T) {
	list := UnorderedList("a").Add(Item("parent", Item("child1"), Item("child2")))
	out := renderDoc(t, NewHTMLDocument(HTMLOptions{}).Add(list))
	if !strings.Contains(out, "<li>parent\n<ul>\n<li>child1</li>\n<li>child2</li>\n</ul>\n</li>") {
		t.Errorf("nested list not rendered as expected, got:\n%s", out)
	}
}

func TestHTMLPerBlockStyle(t *testing.T) {
	doc := NewHTMLDocument(HTMLOptions{}).
		Add(Heading(2, "T").WithStyle(&Style{TextColor: "#FF0000"})).
		Add(Paragraph("p").WithStyle(&Style{Italic: true}))
	out := renderDoc(t, doc)
	if !strings.Contains(out, `<h2 style="color:#FF0000">T</h2>`) {
		t.Errorf("heading style missing, got:\n%s", out)
	}
	if !strings.Contains(out, `<p style="font-style:italic">p</p>`) {
		t.Errorf("paragraph style missing, got:\n%s", out)
	}
}

func TestHTMLTableCaption(t *testing.T) {
	doc := NewHTMLDocument(HTMLOptions{}).
		Add(TableBlock(sampleDocTable()).WithCaption("Q<1> results"))
	out := renderDoc(t, doc)
	if !strings.Contains(out, "<caption>Q&lt;1&gt; results</caption>") {
		t.Errorf("caption missing/unescaped, got:\n%s", out)
	}
}

func TestHTMLTableOfContents(t *testing.T) {
	doc := NewHTMLDocument(HTMLOptions{TableOfContents: true}).
		Heading(2, "Intro").
		Section(3, "Details", Paragraph("x")).
		Heading(2, "Intro") // duplicate -> id dedup

	out := renderDoc(t, doc)

	for _, want := range []string{
		`<nav class="toc">`,
		`<a href="#intro">Intro</a>`,
		`<a href="#details">Details</a>`,
		`<a href="#intro-1">Intro</a>`,
		`<h2 id="intro">Intro</h2>`,
		`<h3 id="details">Details</h3>`,
		`<h2 id="intro-1">Intro</h2>`,
	} {
		if !strings.Contains(out, want) {
			t.Errorf("TOC output missing %q", want)
		}
	}
}

func TestHTMLNoTOCByDefault(t *testing.T) {
	out := renderDoc(t, NewHTMLDocument(HTMLOptions{}).Heading(2, "Intro"))
	if strings.Contains(out, "nav class=\"toc\"") {
		t.Error("TOC should not render when TableOfContents is false")
	}
	if strings.Contains(out, "id=") {
		t.Error("headings should carry no id when TOC is disabled")
	}
}

func TestHTMLTheme(t *testing.T) {
	withTheme := renderDoc(t, NewHTMLDocument(HTMLOptions{Theme: HTMLThemeDefault}).Paragraph("x"))
	if !strings.Contains(withTheme, "tbody tr:nth-child(even)") {
		t.Error("default theme stylesheet not injected")
	}
	if !strings.Contains(withTheme, `<meta name="viewport"`) {
		t.Error("viewport meta missing")
	}
	noTheme := renderDoc(t, NewHTMLDocument(HTMLOptions{}).Paragraph("x"))
	if strings.Contains(noTheme, "nth-child(even)") {
		t.Error("no theme stylesheet should be injected by default")
	}
}
