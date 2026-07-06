// html_document.go - Composable HTML documents.
//
// While ExportHTML renders a single table, an HTMLDocument composes a full document from
// a sequence of blocks: headings, paragraphs, lists, definition lists, blockquotes, code,
// images, tables and grouped sections. All blocks share the same document-level HTMLOptions
// (title, description, theme, page styling) and reuse the table rendering (styles, borders,
// merging) from the HTML backend.

package spit

import (
	"fmt"
	"html"
	"io"
	"strings"
)

// HTMLBlock is a renderable piece of an HTML document.
// Implementations are provided by the package (Heading, Paragraph, lists, tables, sections, …);
// use RawHTML to inject arbitrary markup.
type HTMLBlock interface {
	// renderHTML appends the block's markup to b. opts carries document-level settings.
	renderHTML(b *strings.Builder, opts HTMLOptions) error
}

// tocProvider is implemented by blocks that contribute entries to the table of contents.
// collectTOC assigns a unique anchor id to the block (mutating it) and returns its entries.
type tocProvider interface {
	collectTOC(used map[string]int) []tocEntry
}

// tocEntry is a single heading captured for the table of contents.
type tocEntry struct {
	level int
	text  string
	id    string
}

// HTMLDocument composes an HTML document from an ordered list of blocks.
type HTMLDocument struct {
	Options HTMLOptions // Document-level presentation (title, description, theme, page styling)
	Blocks  []HTMLBlock // Body content, rendered in order
}

// NewHTMLDocument creates an empty document with the given options.
func NewHTMLDocument(opts HTMLOptions) *HTMLDocument {
	return &HTMLDocument{Options: opts}
}

// Add appends one or more blocks to the document.
func (d *HTMLDocument) Add(blocks ...HTMLBlock) *HTMLDocument {
	d.Blocks = append(d.Blocks, blocks...)
	return d
}

// Heading appends a heading (level 1-6) block.
func (d *HTMLDocument) Heading(level int, text string) *HTMLDocument {
	return d.Add(Heading(level, text))
}

// Paragraph appends a paragraph block.
func (d *HTMLDocument) Paragraph(text string) *HTMLDocument {
	return d.Add(Paragraph(text))
}

// UnorderedList appends a bulleted list block.
func (d *HTMLDocument) UnorderedList(items ...string) *HTMLDocument {
	return d.Add(UnorderedList(items...))
}

// OrderedList appends a numbered list block.
func (d *HTMLDocument) OrderedList(items ...string) *HTMLDocument {
	return d.Add(OrderedList(items...))
}

// DefinitionList appends a definition (key/value) list block.
func (d *HTMLDocument) DefinitionList(items ...DefinitionItem) *HTMLDocument {
	return d.Add(DefinitionList(items...))
}

// Blockquote appends a blockquote block.
func (d *HTMLDocument) Blockquote(text string) *HTMLDocument {
	return d.Add(Blockquote(text))
}

// CodeBlock appends a preformatted code block.
func (d *HTMLDocument) CodeBlock(code string) *HTMLDocument {
	return d.Add(CodeBlock(code))
}

// HorizontalRule appends a thematic break (<hr>).
func (d *HTMLDocument) HorizontalRule() *HTMLDocument {
	return d.Add(HorizontalRule())
}

// Image appends a standalone image block.
func (d *HTMLDocument) Image(img Image) *HTMLDocument {
	return d.Add(ImageBlock(img))
}

// Table appends a table block.
func (d *HTMLDocument) Table(t *Table) *HTMLDocument {
	return d.Add(TableBlock(t))
}

// Section appends a semantic <section> grouping a heading (of the given level) with
// nested blocks.
func (d *HTMLDocument) Section(level int, title string, blocks ...HTMLBlock) *HTMLDocument {
	return d.Add(Section(level, title, blocks...))
}

// render serializes the whole document to HTML markup.
func (d *HTMLDocument) render() (string, error) {
	// When a table of contents is requested, assign anchor ids to headings/sections
	// (enabling deep-linking) and gather the entries.
	var toc []tocEntry
	if d.Options.TableOfContents {
		used := make(map[string]int)
		for _, block := range d.Blocks {
			if p, ok := block.(tocProvider); ok {
				toc = append(toc, p.collectTOC(used)...)
			}
		}
	}

	var b strings.Builder
	writeDocumentOpen(&b, d.Options)

	if len(toc) > 0 {
		writeTOC(&b, toc)
	}

	for _, block := range d.Blocks {
		if block == nil {
			continue
		}
		if err := block.renderHTML(&b, d.Options); err != nil {
			return "", err
		}
	}

	writeDocumentClose(&b, d.Options)
	return b.String(), nil
}

// ExportHTMLDocument writes a composed HTML document using the generic file writer.
func ExportHTMLDocument(doc *HTMLDocument, params FileWriteParams) (*FileWriteResult, error) {
	if doc == nil {
		return nil, fmt.Errorf("no document provided")
	}

	if params.Extension == "" {
		params.Extension = FormatHTML.String()
	}

	L().Info("Starting HTML document export to file", String("filename", params.Filename))

	markup, err := doc.render()
	if err != nil {
		L().Error("Failed to render HTML document", Error(err))
		return nil, err
	}

	result, err := params.WriteToFile(func(writer io.Writer) error {
		_, werr := io.WriteString(writer, markup)
		return werr
	})
	if err != nil {
		L().Error("Failed to write HTML document to file", Error(err))
		return nil, err
	}

	L().Info("HTML document export completed", String("filename", params.Filename))
	return result, nil
}

// writeTOC renders a linked table of contents from the collected heading entries.
func writeTOC(b *strings.Builder, entries []tocEntry) {
	minLevel := entries[0].level
	for _, e := range entries {
		if e.level < minLevel {
			minLevel = e.level
		}
	}
	b.WriteString("<nav class=\"toc\">\n<ul>\n")
	for _, e := range entries {
		indent := (e.level - minLevel) * 20
		b.WriteString(fmt.Sprintf("<li style=\"margin-left:%dpx\"><a href=\"#%s\">%s</a></li>\n",
			indent, e.id, html.EscapeString(e.text)))
	}
	b.WriteString("</ul>\n</nav>\n")
}

// ---- Heading ----------------------------------------------------------------

// HeadingBlock renders an <h1>..<h6> element.
type HeadingBlock struct {
	level int
	text  string
	id    string
	style *Style
}

// Heading creates a heading block. Level is clamped to the range 1-6.
func Heading(level int, text string) *HeadingBlock {
	return &HeadingBlock{level: clampHeadingLevel(level), text: text}
}

// WithStyle sets an inline style for the heading.
func (h *HeadingBlock) WithStyle(style *Style) *HeadingBlock {
	h.style = style
	return h
}

func (h *HeadingBlock) collectTOC(used map[string]int) []tocEntry {
	h.id = uniqueSlug(h.text, used)
	return []tocEntry{{level: h.level, text: h.text, id: h.id}}
}

func (h *HeadingBlock) renderHTML(b *strings.Builder, _ HTMLOptions) error {
	b.WriteString(fmt.Sprintf("<h%d%s%s>%s</h%d>\n",
		h.level, idAttr(h.id), inlineStyleAttr(h.style), html.EscapeString(h.text), h.level))
	return nil
}

// ---- Paragraph --------------------------------------------------------------

// ParagraphBlock renders a <p> element.
type ParagraphBlock struct {
	text  string
	style *Style
}

// Paragraph creates a paragraph block.
func Paragraph(text string) *ParagraphBlock {
	return &ParagraphBlock{text: text}
}

// WithStyle sets an inline style for the paragraph.
func (p *ParagraphBlock) WithStyle(style *Style) *ParagraphBlock {
	p.style = style
	return p
}

func (p *ParagraphBlock) renderHTML(b *strings.Builder, _ HTMLOptions) error {
	b.WriteString(fmt.Sprintf("<p%s>%s</p>\n", inlineStyleAttr(p.style), html.EscapeString(p.text)))
	return nil
}

// ---- Lists ------------------------------------------------------------------

// ListItem is one entry of a list, optionally containing nested sub-items.
type ListItem struct {
	Text     string
	Children []ListItem
}

// Item builds a ListItem with optional nested children.
func Item(text string, children ...ListItem) ListItem {
	return ListItem{Text: text, Children: children}
}

// ListBlock renders an ordered (<ol>) or unordered (<ul>) list, possibly nested.
type ListBlock struct {
	ordered bool
	items   []ListItem
	style   *Style
}

// UnorderedList creates a bulleted list block from flat string items.
func UnorderedList(items ...string) *ListBlock {
	return &ListBlock{ordered: false, items: stringsToItems(items)}
}

// OrderedList creates a numbered list block from flat string items.
func OrderedList(items ...string) *ListBlock {
	return &ListBlock{ordered: true, items: stringsToItems(items)}
}

// Add appends rich (possibly nested) items to the list.
func (l *ListBlock) Add(items ...ListItem) *ListBlock {
	l.items = append(l.items, items...)
	return l
}

// WithStyle sets an inline style for the outer list element.
func (l *ListBlock) WithStyle(style *Style) *ListBlock {
	l.style = style
	return l
}

func (l *ListBlock) renderHTML(b *strings.Builder, _ HTMLOptions) error {
	writeList(b, l.ordered, l.items, l.style)
	return nil
}

func writeList(b *strings.Builder, ordered bool, items []ListItem, style *Style) {
	tag := "ul"
	if ordered {
		tag = "ol"
	}
	b.WriteString("<" + tag + inlineStyleAttr(style) + ">\n")
	for _, it := range items {
		b.WriteString("<li>" + html.EscapeString(it.Text))
		if len(it.Children) > 0 {
			b.WriteString("\n")
			writeList(b, ordered, it.Children, nil)
		}
		b.WriteString("</li>\n")
	}
	b.WriteString("</" + tag + ">\n")
}

func stringsToItems(strs []string) []ListItem {
	items := make([]ListItem, 0, len(strs))
	for _, s := range strs {
		items = append(items, ListItem{Text: s})
	}
	return items
}

// ---- Definition list --------------------------------------------------------

// DefinitionItem is a single term/description pair of a definition list.
type DefinitionItem struct {
	Term        string
	Description string
}

// Def builds a DefinitionItem.
func Def(term, description string) DefinitionItem {
	return DefinitionItem{Term: term, Description: description}
}

type htmlDefinitionList struct {
	items []DefinitionItem
}

// DefinitionList creates a definition (key/value) list block, ideal for metadata summaries.
func DefinitionList(items ...DefinitionItem) HTMLBlock {
	return htmlDefinitionList{items: items}
}

func (d htmlDefinitionList) renderHTML(b *strings.Builder, _ HTMLOptions) error {
	b.WriteString("<dl>\n")
	for _, it := range d.items {
		b.WriteString(fmt.Sprintf("<dt>%s</dt>\n<dd>%s</dd>\n",
			html.EscapeString(it.Term), html.EscapeString(it.Description)))
	}
	b.WriteString("</dl>\n")
	return nil
}

// ---- Blockquote -------------------------------------------------------------

// BlockquoteBlock renders a <blockquote> element.
type BlockquoteBlock struct {
	text  string
	style *Style
}

// Blockquote creates a blockquote block.
func Blockquote(text string) *BlockquoteBlock {
	return &BlockquoteBlock{text: text}
}

// WithStyle sets an inline style for the blockquote.
func (q *BlockquoteBlock) WithStyle(style *Style) *BlockquoteBlock {
	q.style = style
	return q
}

func (q *BlockquoteBlock) renderHTML(b *strings.Builder, _ HTMLOptions) error {
	b.WriteString(fmt.Sprintf("<blockquote%s>%s</blockquote>\n",
		inlineStyleAttr(q.style), html.EscapeString(q.text)))
	return nil
}

// ---- Code -------------------------------------------------------------------

type htmlCodeBlock struct {
	code string
}

// CodeBlock creates a preformatted code block (<pre><code>).
func CodeBlock(code string) HTMLBlock {
	return htmlCodeBlock{code: code}
}

func (c htmlCodeBlock) renderHTML(b *strings.Builder, _ HTMLOptions) error {
	b.WriteString("<pre><code>" + html.EscapeString(c.code) + "</code></pre>\n")
	return nil
}

// ---- Horizontal rule --------------------------------------------------------

type htmlRule struct{}

// HorizontalRule creates a thematic break block (<hr>).
func HorizontalRule() HTMLBlock {
	return htmlRule{}
}

func (htmlRule) renderHTML(b *strings.Builder, _ HTMLOptions) error {
	b.WriteString("<hr>\n")
	return nil
}

// ---- Image ------------------------------------------------------------------

type htmlImageBlock struct {
	img Image
}

// ImageBlock creates a standalone image block, reusing the Image cell value type.
func ImageBlock(img Image) HTMLBlock {
	return htmlImageBlock{img: img}
}

func (ib htmlImageBlock) renderHTML(b *strings.Builder, _ HTMLOptions) error {
	b.WriteString(imgTag(ib.img) + "\n")
	return nil
}

// ---- Table ------------------------------------------------------------------

// TableContent renders a Table using the HTML backend, with an optional <caption>.
type TableContent struct {
	table   *Table
	caption string
	style   *Style
}

// TableBlock creates a block that renders a table (with its full styling/merging).
func TableBlock(t *Table) *TableContent {
	return &TableContent{table: t}
}

// WithCaption sets an accessible <caption> for the table.
func (tc *TableContent) WithCaption(caption string) *TableContent {
	tc.caption = caption
	return tc
}

// WithStyle overrides the document TableStyle for this table only.
func (tc *TableContent) WithStyle(style *Style) *TableContent {
	tc.style = style
	return tc
}

func (tc *TableContent) renderHTML(b *strings.Builder, opts HTMLOptions) error {
	if tc.table == nil {
		return nil
	}
	o := opts
	if tc.style != nil {
		o.TableStyle = tc.style
	}
	export := &htmlExport{table: tc.table, opts: o, caption: tc.caption, grid: make(map[int]map[int]*htmlCell)}
	if err := export.build(); err != nil {
		return err
	}
	export.writeTable(b)
	return nil
}

// ---- Section ----------------------------------------------------------------

// SectionBlock renders a semantic <section> with a heading followed by nested blocks.
type SectionBlock struct {
	level  int
	title  string
	id     string
	blocks []HTMLBlock
	style  *Style
}

// Section creates a <section> grouping a heading (of the given level, clamped to 1-6)
// with the provided nested blocks.
func Section(level int, title string, blocks ...HTMLBlock) *SectionBlock {
	return &SectionBlock{level: clampHeadingLevel(level), title: title, blocks: blocks}
}

// Add appends nested blocks to the section.
func (s *SectionBlock) Add(blocks ...HTMLBlock) *SectionBlock {
	s.blocks = append(s.blocks, blocks...)
	return s
}

// WithStyle sets an inline style for the <section> element.
func (s *SectionBlock) WithStyle(style *Style) *SectionBlock {
	s.style = style
	return s
}

func (s *SectionBlock) collectTOC(used map[string]int) []tocEntry {
	entries := []tocEntry{}
	if s.title != "" {
		s.id = uniqueSlug(s.title, used)
		entries = append(entries, tocEntry{level: s.level, text: s.title, id: s.id})
	}
	for _, block := range s.blocks {
		if p, ok := block.(tocProvider); ok {
			entries = append(entries, p.collectTOC(used)...)
		}
	}
	return entries
}

func (s *SectionBlock) renderHTML(b *strings.Builder, opts HTMLOptions) error {
	b.WriteString("<section" + inlineStyleAttr(s.style) + ">\n")
	if s.title != "" {
		b.WriteString(fmt.Sprintf("<h%d%s>%s</h%d>\n",
			s.level, idAttr(s.id), html.EscapeString(s.title), s.level))
	}
	for _, block := range s.blocks {
		if block == nil {
			continue
		}
		if err := block.renderHTML(b, opts); err != nil {
			return err
		}
	}
	b.WriteString("</section>\n")
	return nil
}

// ---- Raw HTML ---------------------------------------------------------------

type htmlRaw struct {
	markup string
}

// RawHTML creates a block that emits the given markup verbatim (not escaped).
// Only pass trusted content.
func RawHTML(markup string) HTMLBlock {
	return htmlRaw{markup: markup}
}

func (r htmlRaw) renderHTML(b *strings.Builder, _ HTMLOptions) error {
	b.WriteString(r.markup)
	if !strings.HasSuffix(r.markup, "\n") {
		b.WriteString("\n")
	}
	return nil
}

// ---- Helpers ----------------------------------------------------------------

// clampHeadingLevel bounds a heading level to the valid HTML range 1-6.
func clampHeadingLevel(level int) int {
	if level < 1 {
		return 1
	}
	if level > 6 {
		return 6
	}
	return level
}

// idAttr returns an ` id="..."` attribute fragment, or "" when id is empty.
func idAttr(id string) string {
	if id == "" {
		return ""
	}
	return fmt.Sprintf(" id=\"%s\"", id)
}

// inlineStyleAttr returns a ` style="..."` attribute fragment for a Style, or "" when empty.
func inlineStyleAttr(style *Style) string {
	css := styleToCSS(style)
	if css == "" {
		return ""
	}
	return fmt.Sprintf(" style=\"%s\"", css)
}

// uniqueSlug builds a URL-friendly, de-duplicated anchor id from text.
func uniqueSlug(text string, used map[string]int) string {
	slug := slugify(text)
	if slug == "" {
		slug = "section"
	}
	count := used[slug]
	used[slug]++
	if count > 0 {
		return fmt.Sprintf("%s-%d", slug, count)
	}
	return slug
}

// slugify lowercases text and replaces runs of non-alphanumeric characters with a single '-'.
func slugify(text string) string {
	var b strings.Builder
	prevDash := false
	for _, r := range strings.ToLower(text) {
		switch {
		case (r >= 'a' && r <= 'z') || (r >= '0' && r <= '9'):
			b.WriteRune(r)
			prevDash = false
		default:
			if !prevDash {
				b.WriteByte('-')
				prevDash = true
			}
		}
	}
	return strings.Trim(b.String(), "-")
}
