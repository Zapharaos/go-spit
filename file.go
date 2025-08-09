package go_spit

import (
	"github.com/Zapharaos/go-spit/internal/file"
)

// FileWriteOptions contains generic options for file writing
type FileWriteOptions struct {
	file.WriteOptions
}

// FileWriteResult contains the result of file writing operation
type FileWriteResult struct {
	file.WriteResult
}

// SanitizeFilename sanitizes a string to be safe for use as a filename
// It's the user's responsibility to handle an empty string after sanitization.
func (fwo FileWriteOptions) SanitizeFilename() string {
	return fwo.SanitizeFilename()
}

// RemoveFile safely removes a file
func (fwr FileWriteResult) RemoveFile() error {
	return fwr.RemoveFile()
}
