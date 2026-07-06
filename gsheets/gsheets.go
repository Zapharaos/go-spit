// Package gsheets exports go-spit tables to Google Sheets.
//
// It is an optional, separately-versioned module so the core go-spit library stays
// dependency-light: only projects that import this package pull in the Google API
// client. This package never handles credentials — the caller builds and passes an
// authenticated *sheets.Service (via a service account, OAuth, or Application Default
// Credentials), which requires the https://www.googleapis.com/auth/spreadsheets scope.
//
// The export reuses go-spit's shared merging and styling pipelines by implementing the
// spit.TableOperations interface and translating the operations into Google Sheets API
// batchUpdate requests.
package gsheets

import (
	"context"
	"errors"
	"fmt"

	spit "github.com/Zapharaos/go-spit"
	"google.golang.org/api/sheets/v4"
)

// Sheet pairs a go-spit table with the name of the target sheet (tab).
type Sheet struct {
	Name  string      // Sheet (tab) name; defaults to "Sheet1" when empty
	Table *spit.Table // Table to render into the sheet
}

// Options configures a Google Sheets export.
type Options struct {
	// Title sets the spreadsheet title. It is only used when creating a new
	// spreadsheet (i.e. when spreadsheetID is empty).
	Title string
}

// Result describes the spreadsheet that was written.
type Result struct {
	SpreadsheetID string // The spreadsheet ID (generated when a new one is created)
	URL           string // A convenience edit URL for the spreadsheet
}

// ExportGoogleSheets writes one or more tables to a Google Sheets spreadsheet.
//
// The caller must provide an authenticated svc. When spreadsheetID is empty a new
// spreadsheet is created (titled opts.Title); otherwise the existing spreadsheet is
// updated, creating any missing tabs. Each table is rendered with its full styling,
// borders and cell merging.
func ExportGoogleSheets(ctx context.Context, svc *sheets.Service, spreadsheetID string, in []Sheet, opts Options) (*Result, error) {
	if svc == nil {
		return nil, errors.New("gsheets: nil *sheets.Service")
	}
	if len(in) == 0 {
		return nil, errors.New("gsheets: no sheets provided")
	}

	// Normalize and validate input.
	sheetsIn := make([]Sheet, len(in))
	for i, s := range in {
		if s.Table == nil {
			return nil, fmt.Errorf("gsheets: nil table for sheet %q", s.Name)
		}
		if s.Name == "" {
			s.Name = "Sheet1"
		}
		sheetsIn[i] = s
	}

	// Create a fresh spreadsheet (with all requested tabs) when no ID is given.
	if spreadsheetID == "" {
		newSS := &sheets.Spreadsheet{
			Properties: &sheets.SpreadsheetProperties{Title: opts.Title},
		}
		for _, s := range sheetsIn {
			newSS.Sheets = append(newSS.Sheets, &sheets.Sheet{
				Properties: &sheets.SheetProperties{Title: s.Name},
			})
		}
		created, err := svc.Spreadsheets.Create(newSS).Context(ctx).Do()
		if err != nil {
			return nil, fmt.Errorf("gsheets: create spreadsheet: %w", err)
		}
		spreadsheetID = created.SpreadsheetId
	}

	// Resolve each target sheet to its numeric ID, adding any missing tabs.
	ids, err := resolveSheetIDs(ctx, svc, spreadsheetID, sheetsIn)
	if err != nil {
		return nil, err
	}

	// Build all cell and merge requests across every table.
	var requests []*sheets.Request
	for _, s := range sheetsIn {
		g := newGSheetTable(s.Table, ids[s.Name])
		if err := g.build(); err != nil {
			return nil, fmt.Errorf("gsheets: build sheet %q: %w", s.Name, err)
		}
		requests = append(requests, g.requests()...)
	}

	// Apply everything in a single batched network call.
	if len(requests) > 0 {
		_, err = svc.Spreadsheets.BatchUpdate(spreadsheetID, &sheets.BatchUpdateSpreadsheetRequest{
			Requests: requests,
		}).Context(ctx).Do()
		if err != nil {
			return nil, fmt.Errorf("gsheets: batch update: %w", err)
		}
	}

	return &Result{
		SpreadsheetID: spreadsheetID,
		URL:           "https://docs.google.com/spreadsheets/d/" + spreadsheetID + "/edit",
	}, nil
}

// resolveSheetIDs returns a map of sheet name to numeric sheet ID for the spreadsheet,
// creating any tabs that do not yet exist.
func resolveSheetIDs(ctx context.Context, svc *sheets.Service, spreadsheetID string, in []Sheet) (map[string]int64, error) {
	ss, err := svc.Spreadsheets.Get(spreadsheetID).Context(ctx).Do()
	if err != nil {
		return nil, fmt.Errorf("gsheets: get spreadsheet: %w", err)
	}
	ids := make(map[string]int64)
	for _, sh := range ss.Sheets {
		ids[sh.Properties.Title] = sh.Properties.SheetId
	}

	// Add any missing tabs, then refresh the ID map.
	var addReqs []*sheets.Request
	for _, s := range in {
		if _, ok := ids[s.Name]; !ok {
			addReqs = append(addReqs, &sheets.Request{AddSheet: &sheets.AddSheetRequest{
				Properties: &sheets.SheetProperties{Title: s.Name},
			}})
		}
	}
	if len(addReqs) > 0 {
		if _, err := svc.Spreadsheets.BatchUpdate(spreadsheetID, &sheets.BatchUpdateSpreadsheetRequest{
			Requests: addReqs,
		}).Context(ctx).Do(); err != nil {
			return nil, fmt.Errorf("gsheets: add sheets: %w", err)
		}
		ss, err = svc.Spreadsheets.Get(spreadsheetID).Context(ctx).Do()
		if err != nil {
			return nil, fmt.Errorf("gsheets: get spreadsheet: %w", err)
		}
		for _, sh := range ss.Sheets {
			ids[sh.Properties.Title] = sh.Properties.SheetId
		}
	}
	return ids, nil
}
