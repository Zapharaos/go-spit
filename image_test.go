package spit

import (
	"bytes"
	"image"
	"image/color"
	"image/png"
	"os"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

// onePixelPNG is a valid 1x1 PNG used to exercise real image insertion.
var onePixelPNG = func() []byte {
	var buf bytes.Buffer
	img := image.NewRGBA(image.Rect(0, 0, 1, 1))
	img.Set(0, 0, color.RGBA{R: 255, A: 255})
	if err := png.Encode(&buf, img); err != nil {
		panic(err)
	}
	return buf.Bytes()
}()

func TestImageHelpers(t *testing.T) {
	if got := NewImageURL("https://x/y.png").TextValue(); got != "https://x/y.png" {
		t.Errorf("TextValue() with URL = %q", got)
	}
	if got := (Image{AltText: "logo"}).TextValue(); got != "logo" {
		t.Errorf("TextValue() alt fallback = %q", got)
	}

	uri := NewImageBytes([]byte("abc"), "image/png").DataURI()
	if !strings.HasPrefix(uri, "data:image/png;base64,") {
		t.Errorf("DataURI prefix = %q", uri)
	}
	if (Image{}).DataURI() != "" {
		t.Error("DataURI() with no bytes should be empty")
	}

	extCases := map[string]string{
		"image/png": ".png", "image/jpeg": ".jpg", "image/svg+xml": ".svg",
		"png": ".png", ".gif": ".gif", "": "",
	}
	for mime, want := range extCases {
		if got := extensionFromMIME(mime); got != want {
			t.Errorf("extensionFromMIME(%q) = %q, want %q", mime, got, want)
		}
	}

	if _, ok := asImage(NewImageURL("u")); !ok {
		t.Error("asImage should accept Image value")
	}
	iv := NewImageURL("u")
	if _, ok := asImage(&iv); !ok {
		t.Error("asImage should accept *Image")
	}
	if _, ok := asImage("not an image"); ok {
		t.Error("asImage should reject non-image values")
	}
}

func TestHTMLImageURL(t *testing.T) {
	data := DataSlice{{"logo": NewImageURL("https://acme.com/logo.png").WithAltText("Acme").WithSize(40, 20)}}
	table := NewTable(data, Columns{NewColumn("logo", "Logo")}, true)
	out := buildHTML(t, table, HTMLOptions{})
	if !strings.Contains(out, `<img src="https://acme.com/logo.png" alt="Acme" width="40" height="20">`) {
		t.Errorf("unexpected img tag in:\n%s", out)
	}
}

func TestHTMLImageBytes(t *testing.T) {
	data := DataSlice{{"logo": NewImageBytes(onePixelPNG, "image/png").WithAltText("dot")}}
	table := NewTable(data, Columns{NewColumn("logo", "Logo")}, true)
	out := buildHTML(t, table, HTMLOptions{})
	if !strings.Contains(out, `<img src="data:image/png;base64,`) {
		t.Errorf("expected data URI img in:\n%s", out)
	}
	if !strings.Contains(out, `alt="dot"`) {
		t.Error("expected alt text on embedded image")
	}
}

func TestCSVImageFallback(t *testing.T) {
	dir := t.TempDir()
	data := DataSlice{{"logo": NewImageURL("https://acme.com/logo.png")}}
	table := NewTable(data, Columns{NewColumn("logo", "Logo")}, true)

	res, err := ExportCSV(",", table, FileWriteParams{Filename: "img", Filepath: dir, OverwriteFile: true})
	if err != nil {
		t.Fatalf("ExportCSV failed: %v", err)
	}
	content, err := os.ReadFile(res.Filepath)
	if err != nil {
		t.Fatalf("read failed: %v", err)
	}
	if !strings.Contains(string(content), "https://acme.com/logo.png") {
		t.Errorf("expected URL text in CSV, got:\n%s", content)
	}
}

func TestXLSXImageFromBytes(t *testing.T) {
	table := NewTable(DataSlice{{"logo": "x"}}, Columns{NewColumn("logo", "Logo")}, true)
	file := excelize.NewFile()
	defer func() { _ = file.Close() }()
	ops := NewTableExcelize("Sheet1", table).WithFile(file)

	if err := ops.SetCellImage(1, 2, NewImageBytes(onePixelPNG, "image/png").WithAltText("dot")); err != nil {
		t.Fatalf("SetCellImage failed: %v", err)
	}

	pics, err := file.GetPictures("Sheet1", "A2")
	if err != nil {
		t.Fatalf("GetPictures failed: %v", err)
	}
	if len(pics) == 0 {
		t.Error("expected a picture anchored at A2")
	}
}

func TestXLSXImageEndToEnd(t *testing.T) {
	dir := t.TempDir()
	data := DataSlice{{"logo": NewImageBytes(onePixelPNG, "image/png")}}
	table := NewTable(data, Columns{NewColumn("logo", "Logo")}, true)
	s := NewSpreadsheetExcelize("Sheet1", table)

	res, err := ExportXLSX(s, FileWriteParams{Filename: "imgx", Filepath: dir, OverwriteFile: true})
	if err != nil {
		t.Fatalf("ExportXLSX failed: %v", err)
	}

	f, err := excelize.OpenFile(res.Filepath)
	if err != nil {
		t.Fatalf("OpenFile failed: %v", err)
	}
	defer func() { _ = f.Close() }()

	// Header occupies row 1, so the single data row's image lands at A2.
	pics, err := f.GetPictures("Sheet1", "A2")
	if err != nil {
		t.Fatalf("GetPictures failed: %v", err)
	}
	if len(pics) == 0 {
		t.Error("expected a picture in the exported workbook at A2")
	}
}
