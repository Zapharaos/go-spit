# File Options

Both `ExportCSV` and `ExportXLSX` accept a `FileWriteParams` value that controls **where** and
**how** the output file is written.

```go
type FileWriteParams struct {
	Filename      string // Desired filename (without extension)
	Filepath      string // Directory to write the file to (used when UseTempFile is false)
	UseTempFile   bool   // Optional: create a temp file (default: false)
	UseGzip       bool   // Optional: compress the output with gzip
	OverwriteFile bool   // Optional: overwrite an existing file (default: false)
	Extension     string // File extension (e.g. "csv", "xlsx"); set automatically when empty
}
```

## Fields

| Field           | Description                                                                                   |
|-----------------|-----------------------------------------------------------------------------------------------|
| `Filename`      | The base name, without extension. It is **sanitized** before use (see below).                 |
| `Filepath`      | Target directory. Empty or `"."` means the current directory. Missing directories are created. |
| `UseTempFile`   | When `true`, a uniquely named temp file is created in `Filepath` instead of a fixed name.     |
| `UseGzip`       | When `true`, the output is gzip-compressed and a `.gz` suffix is appended.                     |
| `OverwriteFile` | When `false` (default), exporting fails if the target file already exists.                     |
| `Extension`     | Normally left empty so the exporter sets `csv`/`xlsx` automatically.                           |

## Example

```go
params := spit.FileWriteParams{
	Filename:      "report",            // becomes report.csv / report.xlsx
	Filepath:      "/path/to/output",   // created if it does not exist
	UseTempFile:   false,
	UseGzip:       true,                // produces report.csv.gz
	OverwriteFile: true,
}

result, err := spit.ExportCSV(",", table, params)
```

## The result

On success, the exporter returns a `FileWriteResult`:

```go
type FileWriteResult struct {
	Filepath string // Full path to the created file
	Filename string // Final filename (including extension and any modifications)
}
```

Use `result.Filepath` to locate the file. When you no longer need it, remove it with
`RemoveFile`, which safely handles missing files:

```go
defer func() {
	if err := result.RemoveFile(); err != nil {
		log.Printf("failed to remove file: %v", err)
	}
}()
```

## Filename sanitization

Filenames are sanitized with `SanitizeFilename` before the file is created. This:

- Transliterates accented characters (`é` → `e`, `ü` → `u`, `ç` → `c`, …).
- Replaces spaces and filesystem-unsafe characters (`/ \ : * ? " < > |`) with underscores.
- Collapses consecutive underscores and trims leading/trailing underscores.

If a filename becomes empty after sanitization, the export returns an error. You can call
`SanitizeFilename` yourself to preview the result:

```go
clean := spit.SanitizeFilename("Q1 Report: 2024/Final") // "Q1_Report_2024_Final"
```

## Temporary files

Set `UseTempFile: true` to write to a uniquely named temporary file. The file is created in
`Filepath` (or the OS temp location when `Filepath` is empty), using a pattern derived from
`Filename`. This is useful when you want to stream or move the file afterwards and avoid clobbering
existing output.

## Gzip compression

Set `UseGzip: true` to compress the output. The exporter appends `.gz` to the filename and writes
the data through a gzip stream, so `report.csv` becomes `report.csv.gz`.
