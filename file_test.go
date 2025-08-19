package spit

import (
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestFileWriteParams_SanitizeFilename(t *testing.T) {
	tests := []struct {
		name     string
		input    string
		expected string
	}{
		{"Simple", "file", "file"},
		{"Spaces", "my file", "my_file"},
		{"Accents", "Résumé", "Resume"},
		{"SpecialChars", "a/b\\c:d*e?f\"g<h>i|j", "a_b_c_d_e_f_g_h_i_j"},
		{"ConsecutiveUnderscores", "a__b", "a_b"},
		{"TrimUnderscores", "_abc_", "abc"},
		{"Empty", "", ""},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			fwp := FileWriteParams{Filename: tt.input}
			got := fwp.SanitizeFilename()
			if got != tt.expected {
				t.Errorf("SanitizeFilename(%q) = %q, want %q", tt.input, got, tt.expected)
			}
		})
	}
}

func TestFileWriteParams_writeToFile(t *testing.T) {
	tmpDir := t.TempDir()
	tests := []struct {
		name       string
		params     FileWriteParams
		writeFunc  func(io.Writer) error
		expectErr  bool
		expectGzip bool
		overwrite  bool
	}{
		{
			name: "Regular file write",
			params: FileWriteParams{
				Filename:      "testfile",
				Filepath:      tmpDir,
				extension:     "txt",
				UseTempFile:   false,
				UseGzip:       false,
				OverwriteFile: true,
			},
			writeFunc: func(w io.Writer) error {
				_, err := w.Write([]byte("hello"))
				return err
			},
			expectErr:  false,
			expectGzip: false,
			overwrite:  true,
		},
		{
			name: "Gzip file write",
			params: FileWriteParams{
				Filename:      "gzfile",
				Filepath:      tmpDir,
				extension:     "log",
				UseTempFile:   false,
				UseGzip:       true,
				OverwriteFile: true,
			},
			writeFunc: func(w io.Writer) error {
				_, err := w.Write([]byte("gzipdata"))
				return err
			},
			expectErr:  false,
			expectGzip: true,
			overwrite:  true,
		},
		{
			name: "Temp file write",
			params: FileWriteParams{
				Filename:      "tempfile",
				Filepath:      tmpDir,
				extension:     "tmp",
				UseTempFile:   true,
				UseGzip:       false,
				OverwriteFile: false,
			},
			writeFunc: func(w io.Writer) error {
				_, err := w.Write([]byte("tempdata"))
				return err
			},
			expectErr:  false,
			expectGzip: false,
			overwrite:  false,
		},
		{
			name: "Empty filename after sanitization",
			params: FileWriteParams{
				Filename:      "",
				Filepath:      tmpDir,
				extension:     "txt",
				UseTempFile:   false,
				UseGzip:       false,
				OverwriteFile: false,
			},
			writeFunc: func(w io.Writer) error {
				return nil
			},
			expectErr: true,
		},
		{
			name: "File already exists and overwrite is false",
			params: FileWriteParams{
				Filename:      "existsfile",
				Filepath:      tmpDir,
				extension:     "txt",
				UseTempFile:   false,
				UseGzip:       false,
				OverwriteFile: false,
			},
			writeFunc: func(w io.Writer) error {
				_, err := w.Write([]byte("data"))
				return err
			},
			expectErr: true,
		},
	}
	// Pre-create a file for the "already exists" test
	existsPath := filepath.Join(tmpDir, "existsfile.txt")
	os.WriteFile(existsPath, []byte("exists"), 0644)

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			res, err := tt.params.writeToFile(tt.writeFunc)
			if tt.expectErr {
				if err == nil {
					t.Errorf("expected error, got nil")
				}
				return
			}
			if err != nil {
				t.Fatalf("unexpected error: %v", err)
			}
			if res == nil || !strings.Contains(res.Filename, tt.params.Filename) {
				t.Errorf("unexpected result: %+v", res)
			}
			// Check file existence
			if _, err := os.Stat(res.Filepath); err != nil {
				t.Errorf("file not created: %v", err)
			}
			// Optionally check for gzip extension
			if tt.expectGzip && !strings.HasSuffix(res.Filename, ".gz") {
				t.Errorf("expected gzip file, got %s", res.Filename)
			}
		})
	}
}

func TestFileWriteResult_RemoveFile(t *testing.T) {
	tmpDir := t.TempDir()
	filePath := filepath.Join(tmpDir, "toremove.txt")
	os.WriteFile(filePath, []byte("data"), 0644)

	tests := []struct {
		name    string
		result  FileWriteResult
		setup   func()
		wantErr bool
	}{
		{
			name:    "Remove existing file",
			result:  FileWriteResult{Filepath: filePath, Filename: "toremove.txt"},
			setup:   func() {},
			wantErr: false,
		},
		{
			name:    "Remove non-existent file",
			result:  FileWriteResult{Filepath: filepath.Join(tmpDir, "nofile.txt"), Filename: "nofile.txt"},
			setup:   func() {},
			wantErr: false,
		},
		{
			name:    "Empty file path",
			result:  FileWriteResult{Filepath: "", Filename: ""},
			setup:   func() {},
			wantErr: false,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			tt.setup()
			err := tt.result.RemoveFile()
			if tt.wantErr && err == nil {
				t.Errorf("expected error, got nil")
			}
			if !tt.wantErr && err != nil {
				t.Errorf("unexpected error: %v", err)
			}
		})
	}
}
