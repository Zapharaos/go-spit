// image.go - Image cell values.
//
// This file defines the Image type used to place images into table cells. An Image
// can reference an external/local source (URL) or carry embedded binary content
// (Bytes + MIME). Backends render it in a format-appropriate way:
//   - HTML: an <img> element (data URI when embedded, otherwise src=URL)
//   - XLSX: a cell-anchored picture (via Excelize AddPicture/AddPictureFromBytes)
//   - CSV: a textual fallback (the URL, or the alt text when no URL is set)

package spit

import (
	"encoding/base64"
	"strings"
)

// Image represents an image value for a table cell.
//
// Provide either a reference via URL or embedded content via Bytes together with MIME:
//   - URL: a remote URL (HTML) or a local file path (XLSX). Remote URLs are not fetched
//     for XLSX; use Bytes to embed remote images in spreadsheets.
//   - Bytes + MIME: embedded content, rendered as a data URI in HTML and inserted from
//     bytes in XLSX. MIME must be set (e.g. "image/png") so the format can be resolved.
//
// Width and Height are optional hints applied to HTML output only.
type Image struct {
	URL     string // Remote URL (HTML) or local file path (XLSX)
	Bytes   []byte // Optional embedded binary content
	MIME    string // MIME type for embedded content (e.g. "image/png"); required with Bytes
	AltText string // Alternative text (accessibility / CSV fallback)
	Width   int    // Optional width in pixels (HTML only)
	Height  int    // Optional height in pixels (HTML only)
}

// NewImageURL creates an Image that references an external URL or local file path.
func NewImageURL(url string) Image {
	return Image{URL: url}
}

// NewImageBytes creates an Image with embedded content of the given MIME type.
func NewImageBytes(data []byte, mime string) Image {
	return Image{Bytes: data, MIME: mime}
}

// WithAltText sets the alternative text for this image.
func (img Image) WithAltText(alt string) Image {
	img.AltText = alt
	return img
}

// WithSize sets the width and height (in pixels) for this image.
func (img Image) WithSize(width, height int) Image {
	img.Width = width
	img.Height = height
	return img
}

// HasBytes reports whether the image carries embedded binary content.
func (img Image) HasBytes() bool {
	return len(img.Bytes) > 0
}

// DataURI returns a base64-encoded data URI for the embedded content, or an empty
// string when no bytes are present.
func (img Image) DataURI() string {
	if !img.HasBytes() {
		return ""
	}
	mime := img.MIME
	if mime == "" {
		mime = "application/octet-stream"
	}
	return "data:" + mime + ";base64," + base64.StdEncoding.EncodeToString(img.Bytes)
}

// TextValue returns the textual fallback for the image: the URL when set, otherwise
// the alt text. Used by text-only backends such as CSV.
func (img Image) TextValue() string {
	if img.URL != "" {
		return img.URL
	}
	return img.AltText
}

// extensionFromMIME maps a MIME type (or an already-extension-like string) to the file
// extension Excelize expects (including the leading dot). Returns "" when unknown.
func extensionFromMIME(mime string) string {
	m := strings.ToLower(strings.TrimSpace(mime))
	switch m {
	case "image/png":
		return ".png"
	case "image/jpeg", "image/jpg":
		return ".jpg"
	case "image/gif":
		return ".gif"
	case "image/bmp":
		return ".bmp"
	case "image/tiff":
		return ".tif"
	case "image/svg+xml":
		return ".svg"
	case "image/webp":
		return ".webp"
	}
	// Allow callers to pass an extension directly (e.g. ".png" or "png").
	if strings.HasPrefix(m, ".") {
		return m
	}
	if m != "" && !strings.Contains(m, "/") {
		return "." + m
	}
	return ""
}

// asImage extracts an Image from a cell value, accepting both Image and *Image.
func asImage(value interface{}) (Image, bool) {
	switch v := value.(type) {
	case Image:
		return v, true
	case *Image:
		if v != nil {
			return *v, true
		}
	}
	return Image{}, false
}
