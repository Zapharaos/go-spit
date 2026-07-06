package gsheets

import (
	"context"
	"testing"

	spit "github.com/Zapharaos/go-spit"
	"google.golang.org/api/sheets/v4"
)

// findUpdateCells returns the first UpdateCells request, or nil.
func findUpdateCells(reqs []*sheets.Request) *sheets.UpdateCellsRequest {
	for _, r := range reqs {
		if r.UpdateCells != nil {
			return r.UpdateCells
		}
	}
	return nil
}

func TestBuildValuesTypesAndHeaderStyle(t *testing.T) {
	data := spit.DataSlice{{"name": "John", "age": 30}}
	cols := spit.Columns{spit.NewColumn("name", "Name"), spit.NewColumn("age", "Age")}
	g := newGSheetTable(spit.NewTable(data, cols, true), 7)
	if err := g.build(); err != nil {
		t.Fatalf("build: %v", err)
	}

	uc := findUpdateCells(g.requests())
	if uc == nil {
		t.Fatal("expected an UpdateCells request")
	}
	if uc.Start.SheetId != 7 {
		t.Errorf("sheet id = %d, want 7", uc.Start.SheetId)
	}
	if len(uc.Rows) != 2 {
		t.Fatalf("rows = %d, want 2 (header + data)", len(uc.Rows))
	}

	// Header cell "Name" is a string and carries the default header style (bold).
	head := uc.Rows[0].Values[0]
	if head.UserEnteredValue == nil || head.UserEnteredValue.StringValue == nil || *head.UserEnteredValue.StringValue != "Name" {
		t.Errorf("header cell value = %+v, want string \"Name\"", head.UserEnteredValue)
	}
	if head.UserEnteredFormat == nil || head.UserEnteredFormat.TextFormat == nil || !head.UserEnteredFormat.TextFormat.Bold {
		t.Error("expected bold default header style")
	}

	// Data row: string "John" and native number 30.
	row := uc.Rows[1].Values
	if row[0].UserEnteredValue.StringValue == nil || *row[0].UserEnteredValue.StringValue != "John" {
		t.Errorf("name cell = %+v, want string \"John\"", row[0].UserEnteredValue)
	}
	if row[1].UserEnteredValue.NumberValue == nil || *row[1].UserEnteredValue.NumberValue != 30 {
		t.Errorf("age cell = %+v, want number 30", row[1].UserEnteredValue)
	}
}

func TestVerticalMergeRequest(t *testing.T) {
	data := spit.DataSlice{{"dept": "Eng"}, {"dept": "Eng"}, {"dept": "Sales"}}
	cols := spit.Columns{
		spit.NewColumn("dept", "Dept").
			WithMerge(spit.NewMergeRules(spit.MergeConditions{spit.MergeConditionIdentical}, nil)),
	}
	g := newGSheetTable(spit.NewTable(data, cols, true), 0)
	if err := g.build(); err != nil {
		t.Fatalf("build: %v", err)
	}

	var merged *sheets.GridRange
	for _, r := range g.requests() {
		if r.MergeCells != nil {
			merged = r.MergeCells.Range
		}
	}
	if merged == nil {
		t.Fatal("expected a MergeCells request for the two identical 'Eng' rows")
	}
	// Header at row 1; the two 'Eng' data rows are sheet rows 2-3 => 0-based A2:A3.
	if merged.StartRowIndex != 1 || merged.EndRowIndex != 3 ||
		merged.StartColumnIndex != 0 || merged.EndColumnIndex != 1 {
		t.Errorf("merge range = %+v, want rows[1,3) cols[0,1)", merged)
	}
}

func TestHyperlinkAndFormula(t *testing.T) {
	data := spit.DataSlice{{"site": "https://x.dev", "sum": "=1+1"}}
	cols := spit.Columns{
		spit.NewColumn("site", "Site").WithFormat(spit.ExcelizeFormatHyperlink),
		spit.NewColumn("sum", "Sum").WithFormat(spit.ExcelizeFormatFormula),
	}
	g := newGSheetTable(spit.NewTable(data, cols, true), 0)
	if err := g.build(); err != nil {
		t.Fatalf("build: %v", err)
	}
	uc := findUpdateCells(g.requests())
	row := uc.Rows[1].Values
	if row[0].UserEnteredValue.FormulaValue == nil || *row[0].UserEnteredValue.FormulaValue != `=HYPERLINK("https://x.dev")` {
		t.Errorf("hyperlink cell = %+v", row[0].UserEnteredValue)
	}
	if row[1].UserEnteredValue.FormulaValue == nil || *row[1].UserEnteredValue.FormulaValue != "=1+1" {
		t.Errorf("formula cell = %+v", row[1].UserEnteredValue)
	}
}

func TestImageBlockFormula(t *testing.T) {
	data := spit.DataSlice{{"logo": spit.NewImageURL("https://acme.com/l.png")}}
	cols := spit.Columns{spit.NewColumn("logo", "Logo")}
	g := newGSheetTable(spit.NewTable(data, cols, true), 0)
	if err := g.build(); err != nil {
		t.Fatalf("build: %v", err)
	}
	uc := findUpdateCells(g.requests())
	cell := uc.Rows[1].Values[0]
	if cell.UserEnteredValue.FormulaValue == nil || *cell.UserEnteredValue.FormulaValue != `=IMAGE("https://acme.com/l.png")` {
		t.Errorf("image cell = %+v, want =IMAGE(...)", cell.UserEnteredValue)
	}
}

func TestExportGoogleSheetsValidation(t *testing.T) {
	ctx := context.Background()
	tbl := spit.NewTable(spit.DataSlice{{"a": 1}}, spit.Columns{spit.NewColumn("a", "A")}, true)

	if _, err := ExportGoogleSheets(ctx, nil, "", []Sheet{{Name: "S", Table: tbl}}, Options{}); err == nil {
		t.Error("expected error for nil service")
	}
	if _, err := ExportGoogleSheets(ctx, &sheets.Service{}, "", nil, Options{}); err == nil {
		t.Error("expected error for no sheets")
	}
	if _, err := ExportGoogleSheets(ctx, &sheets.Service{}, "", []Sheet{{Name: "S", Table: nil}}, Options{}); err == nil {
		t.Error("expected error for nil table")
	}
}
