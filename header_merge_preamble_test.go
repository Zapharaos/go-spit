package spit

import (
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

// makePreambleHeaderTable builds a table that combines a preamble row with a
// two-level hierarchical header. This is the combination that previously caused
// header merges to be applied to the preamble rows instead of the header rows.
func makePreambleHeaderTable() *Table {
	data := DataSlice{
		{"name": "John", "age": 30, "dept": "Eng"},
		{"name": "Jane", "age": 28, "dept": "Eng"},
	}
	return NewTable(data, Columns{
		NewColumn("name", "Name"), // top-level leaf -> vertical header merge across both header rows
		NewColumn("", "Details").WithSubColumns(Columns{
			NewColumn("age", "Age"),
			NewColumn("dept", "Department"),
		}),
	}, true).WithPreamble(PreambleRows{NewPreambleRow("Report 2026")})
}

// TestHeaderMergingWithPreamble is a regression test for header merging offset by
// preamble rows: header merges must target the actual header rows (starting at
// GetHeaderStartRow()), never the preamble rows above them.
func TestHeaderMergingWithPreamble(t *testing.T) {
	t.Run("HTML backend keeps preamble and header labels intact", func(t *testing.T) {
		table := makePreambleHeaderTable()
		export := &htmlExport{table: table, grid: make(map[int]map[int]*htmlCell)}
		if err := export.build(); err != nil {
			t.Fatalf("build failed: %v", err)
		}

		// The preamble sits at grid row 1 and must remain a standalone cell.
		preamble := export.peek(1, 1)
		if preamble == nil {
			t.Fatal("preamble cell missing at (1,1)")
		}
		if preamble.covered || preamble.rowspan != 1 {
			t.Errorf("preamble cell was merged into the header: covered=%v rowspan=%d", preamble.covered, preamble.rowspan)
		}
		if preamble.value != "Report 2026" {
			t.Errorf("preamble value corrupted: %q", preamble.value)
		}

		// The top-level leaf header "Name" lives at the first header row and spans
		// both header rows (rowspan=2), not starting from the preamble row.
		headerStart := table.GetHeaderStartRow() // 2 (one preamble row)
		nameHdr := export.peek(1, headerStart)
		if nameHdr == nil || nameHdr.value != "Name" {
			t.Fatalf("expected 'Name' header at row %d, got %+v", headerStart, nameHdr)
		}
		if nameHdr.rowspan != 2 {
			t.Errorf("expected 'Name' header rowspan=2, got %d", nameHdr.rowspan)
		}

		if out := export.render(); !strings.Contains(out, ">Name<") {
			t.Errorf("top-level leaf header label 'Name' missing from output:\n%s", out)
		}
	})

	t.Run("XLSX backend merges the correct header rows", func(t *testing.T) {
		table := makePreambleHeaderTable()
		file := excelize.NewFile()
		defer func() { _ = file.Close() }()
		ops := NewTableExcelize("Sheet1", table).WithFile(file)

		if err := table.ProcessMerging(ops); err != nil {
			t.Fatalf("ProcessMerging failed: %v", err)
		}

		merges, err := file.GetMergeCells("Sheet1")
		if err != nil {
			t.Fatalf("GetMergeCells failed: %v", err)
		}

		// No merged range may start in the preamble row (row 1).
		for _, m := range merges {
			_, startRow, err := excelize.CellNameToCoordinates(m.GetStartAxis())
			if err != nil {
				t.Fatalf("bad axis %q: %v", m.GetStartAxis(), err)
			}
			if startRow < 2 {
				t.Errorf("merge %s:%s starts in preamble row %d; headers start at row 2",
					m.GetStartAxis(), m.GetEndAxis(), startRow)
			}
		}

		// Expect the two header merges: 'Name' vertical (A2:A3) and 'Details' horizontal (B2:C2).
		want := map[string]bool{"A2:A3": false, "B2:C2": false}
		for _, m := range merges {
			key := m.GetStartAxis() + ":" + m.GetEndAxis()
			if _, ok := want[key]; ok {
				want[key] = true
			}
		}
		for key, found := range want {
			if !found {
				t.Errorf("expected header merge %s, got merges %v", key, merges)
			}
		}
	})
}
