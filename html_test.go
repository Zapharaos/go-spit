package spit

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// buildHTML is a small helper that builds the in-memory export and returns the markup.
func buildHTML(t *testing.T, table *Table, opts HTMLOptions) string {
	t.Helper()
	export := &htmlExport{table: table, opts: opts, grid: make(map[int]map[int]*htmlCell)}
	if err := export.build(); err != nil {
		t.Fatalf("build failed: %v", err)
	}
	return export.render()
}

func TestExportHTMLWritesFile(t *testing.T) {
	dir := t.TempDir()
	table := NewTable(testData, Columns{
		NewColumn("name", "Name"),
		NewColumn("age", "Age"),
		NewColumn("city", "City"),
	}, true)

	result, err := ExportHTML(table, HTMLOptions{Title: "People"}, FileWriteParams{
		Filename:      "people",
		Filepath:      dir,
		OverwriteFile: true,
	})
	if err != nil {
		t.Fatalf("ExportHTML failed: %v", err)
	}
	if filepath.Ext(result.Filename) != ".html" {
		t.Errorf("expected .html extension, got %q", result.Filename)
	}

	content, err := os.ReadFile(result.Filepath)
	if err != nil {
		t.Fatalf("failed to read output: %v", err)
	}
	out := string(content)
	for _, want := range []string{"<!DOCTYPE html>", "<h1>People</h1>", "<th", "John", "Jane"} {
		if !strings.Contains(out, want) {
			t.Errorf("output missing %q", want)
		}
	}
}

func TestExportHTMLNilTable(t *testing.T) {
	if _, err := ExportHTML(nil, HTMLOptions{}, FileWriteParams{Filename: "x"}); err == nil {
		t.Error("expected error for nil table")
	}
}

func TestHTMLEscaping(t *testing.T) {
	data := DataSlice{{"c": "<script>alert(1)</script> & \"quotes\""}}
	table := NewTable(data, Columns{NewColumn("c", "A<b>")}, true)
	out := buildHTML(t, table, HTMLOptions{Title: "T<i>", Description: "d&d"})

	for _, want := range []string{
		"&lt;script&gt;",
		"A&lt;b&gt;",
		"<h1>T&lt;i&gt;</h1>",
		"<p>d&amp;d</p>",
	} {
		if !strings.Contains(out, want) {
			t.Errorf("expected escaped %q in output", want)
		}
	}
	if strings.Contains(out, "<script>alert") {
		t.Error("unescaped script tag leaked into output")
	}
}

func TestHTMLHyperlink(t *testing.T) {
	data := DataSlice{{"site": "https://example.com"}}
	table := NewTable(data, Columns{NewColumn("site", "Site").WithFormat(ExcelizeFormatHyperlink)}, true)
	out := buildHTML(t, table, HTMLOptions{})
	if !strings.Contains(out, `<a href="https://example.com">https://example.com</a>`) {
		t.Errorf("expected hyperlink anchor, got:\n%s", out)
	}
}

func TestHTMLVerticalMerge(t *testing.T) {
	data := DataSlice{
		{"dept": "Eng", "name": "A"},
		{"dept": "Eng", "name": "B"},
		{"dept": "Sales", "name": "C"},
	}
	table := NewTable(data, Columns{
		NewColumn("dept", "Dept").WithMerge(NewMergeRules(MergeConditions{MergeConditionIdentical}, nil)),
		NewColumn("name", "Name"),
	}, true)
	out := buildHTML(t, table, HTMLOptions{})
	if !strings.Contains(out, `rowspan="2"`) {
		t.Errorf("expected rowspan=2 for merged Eng cells, got:\n%s", out)
	}
	// The two "Eng" rows collapse into one rendered cell; only one "Eng" text remains.
	if got := strings.Count(out, ">Eng<"); got != 1 {
		t.Errorf("expected exactly 1 rendered Eng cell, got %d", got)
	}
}

func TestHTMLMultiLevelHeader(t *testing.T) {
	data := DataSlice{{"age": 30, "dept": "Eng"}}
	table := NewTable(data, Columns{
		NewColumn("", "Details").WithSubColumns(Columns{
			NewColumn("age", "Age"),
			NewColumn("dept", "Department"),
		}),
	}, true)
	out := buildHTML(t, table, HTMLOptions{})
	if !strings.Contains(out, `colspan="2"`) {
		t.Errorf("expected colspan=2 for Details header, got:\n%s", out)
	}
}

func TestHTMLFragmentOnly(t *testing.T) {
	table := NewTable(testData, Columns{NewColumn("name", "Name")}, true)
	out := buildHTML(t, table, HTMLOptions{Title: "T", FragmentOnly: true})
	if strings.Contains(out, "<!DOCTYPE") || strings.Contains(out, "<html") || strings.Contains(out, "<body") {
		t.Errorf("fragment should not contain document wrappers, got:\n%s", out)
	}
	if !strings.Contains(out, "<h1>T</h1>") || !strings.Contains(out, "<table") {
		t.Error("fragment should still contain title and table")
	}
}

func TestHTMLBodyAndTableStyle(t *testing.T) {
	table := NewTable(testData, Columns{NewColumn("name", "Name")}, true)
	out := buildHTML(t, table, HTMLOptions{
		BodyStyle:  &Style{FontFamily: "Segoe UI", BackgroundColor: "FAFAFA"},
		TableStyle: &Style{FontSize: 11},
		CustomCSS:  "h1 { color: red; }",
	})
	if !strings.Contains(out, `<body style="background-color:#FAFAFA;font-family:'Segoe UI'">`) {
		t.Errorf("expected body style, got:\n%s", out)
	}
	if !strings.Contains(out, "border-collapse:collapse;font-size:11pt") {
		t.Error("expected table style merged with border-collapse")
	}
	if !strings.Contains(out, "<style>h1 { color: red; }</style>") {
		t.Error("expected custom CSS block")
	}
}

func TestStyleToCSS(t *testing.T) {
	tests := []struct {
		name  string
		style *Style
		want  string
	}{
		{"nil", nil, ""},
		{"bold", &Style{Bold: true}, "font-weight:bold"},
		{"hex without hash", &Style{TextColor: "FF0000"}, "color:#FF0000"},
		{"hex with hash", &Style{TextColor: "#FF0000"}, "color:#FF0000"},
		{"named color", &Style{TextColor: "red"}, "color:red"},
		{"font size", &Style{FontSize: 12}, "font-size:12pt"},
		{"font family with space", &Style{FontFamily: "Times New Roman"}, "font-family:'Times New Roman'"},
		{"align center middle", &Style{Alignment: AlignmentCenterMiddle}, "text-align:center;vertical-align:middle"},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			if got := styleToCSS(tc.style); got != tc.want {
				t.Errorf("styleToCSS() = %q, want %q", got, tc.want)
			}
		})
	}
}

func TestBordersToCSS(t *testing.T) {
	b := Borders{
		Left:   NewBorder(BorderStyleThin),
		Bottom: NewBorder(BorderStyleDouble),
		Top:    NewBorder(BorderStyleNone), // should be skipped
	}
	got := bordersToCSS(b)
	if !strings.Contains(got, "border-left:1px solid #000000") {
		t.Errorf("missing left border in %q", got)
	}
	if !strings.Contains(got, "border-bottom:3px double #000000") {
		t.Errorf("missing bottom border in %q", got)
	}
	if strings.Contains(got, "border-top") {
		t.Errorf("BorderStyleNone should be skipped, got %q", got)
	}
}

func TestGetColumnLetter(t *testing.T) {
	h := &htmlExport{}
	tests := map[int]string{1: "A", 2: "B", 26: "Z", 27: "AA", 28: "AB", 52: "AZ", 53: "BA"}
	for in, want := range tests {
		if got := h.GetColumnLetter(in); got != want {
			t.Errorf("GetColumnLetter(%d) = %q, want %q", in, got, want)
		}
	}
}
