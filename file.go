// file.go - Generic file writing and management.
//
// This file provides utilities for writing data to files with options and managing it.

package spit

import (
	"compress/gzip"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
	"unicode"
)

// FileWriteParams contains generic parameters for file writing
type FileWriteParams struct {
	Filename      string // Desired filename (without extension)
	Filepath      string // Directory to write file to (used if UseTempFile is false)
	UseTempFile   bool   // Optional: use temp file (default: false)
	UseGzip       bool   // Optional: compress with gzip
	OverwriteFile bool   // Optional: overwrite existing file (default: false)
	extension     string // File extension (e.g., ".csv", ".json")
}

// FileWriteResult contains the result of file writing operation
type FileWriteResult struct {
	Filepath string // Full path to the created file
	Filename string // Final filename (including any modifications)
}

// SanitizeFilename sanitizes a string to be safe for use as a filename.
func (fwo FileWriteParams) SanitizeFilename() string {
	// First handle accented characters and special Unicode characters
	result := strings.Map(func(r rune) rune {
		switch r {
		// Accented A variants
		case 'À', 'Á', 'Â', 'Ã', 'Ä', 'Å', 'à', 'á', 'â', 'ã', 'ä', 'å':
			return 'a'
		// Accented E variants
		case 'È', 'É', 'Ê', 'Ë', 'è', 'é', 'ê', 'ë':
			return 'e'
		// Accented I variants
		case 'Ì', 'Í', 'Î', 'Ï', 'ì', 'í', 'î', 'ï':
			return 'i'
		// Accented O variants
		case 'Ò', 'Ó', 'Ô', 'Õ', 'Ö', 'Ø', 'ò', 'ó', 'ô', 'õ', 'ö', 'ø':
			return 'o'
		// Accented U variants
		case 'Ù', 'Ú', 'Û', 'Ü', 'ù', 'ú', 'û', 'ü':
			return 'u'
		// Accented Y variants
		case 'Ý', 'Ÿ', 'ý', 'ÿ':
			return 'y'
		// Accented C variants
		case 'Ç', 'ç':
			return 'c'
		// Accented N variants
		case 'Ñ', 'ñ':
			return 'n'
		// Other special characters
		case 'ß':
			return 's'
		case 'Æ', 'æ':
			return 'a'
		case 'Œ', 'œ':
			return 'o'
		default:
			// Keep the character if it's alphanumeric
			if unicode.IsLetter(r) || unicode.IsDigit(r) {
				return r
			}
			// Replace other characters with underscore
			return '_'
		}
	}, fwo.Filename)

	// Now handle filesystem-problematic characters specifically
	// (in case any slipped through or were originally ASCII)
	result = strings.ReplaceAll(result, " ", "_")
	result = strings.ReplaceAll(result, "/", "_")
	result = strings.ReplaceAll(result, "\\", "_")
	result = strings.ReplaceAll(result, ":", "_")
	result = strings.ReplaceAll(result, "*", "_")
	result = strings.ReplaceAll(result, "?", "_")
	result = strings.ReplaceAll(result, "\"", "_")
	result = strings.ReplaceAll(result, "<", "_")
	result = strings.ReplaceAll(result, ">", "_")
	result = strings.ReplaceAll(result, "|", "_")

	// Remove consecutive underscores and trim
	for strings.Contains(result, "__") {
		result = strings.ReplaceAll(result, "__", "_")
	}
	result = strings.Trim(result, "_")

	// User's responsibility to ensure the filename is not empty after sanitization
	return result
}

// writeToFile writes data to a file with generic options and returns file info
func (fwo FileWriteParams) writeToFile(writeFunc func(io.Writer) error) (*FileWriteResult, error) {
	// Sanitize the filename to ensure it's safe for use
	fwo.Filename = fwo.SanitizeFilename()

	if fwo.Filename == "" {
		return nil, fmt.Errorf("filename is empty after sanitization")
	}

	// Construct filename with extension
	extension := "." + fwo.extension
	fileName := fwo.Filename + extension
	tempFilePattern := fwo.Filename + "_*" + extension

	if fwo.UseGzip {
		fileName += ".gz"
		tempFilePattern += ".gz"
	}

	var filePath string
	var file *os.File
	var err error

	if fwo.UseTempFile {
		L().Debug("creating temp file", String("pattern", tempFilePattern))
		file, err = os.CreateTemp(fwo.Filepath, tempFilePattern)
		if err != nil {
			return nil, fmt.Errorf("failed to create temp file: %w", err)
		}
		filePath = file.Name()
	} else {
		// Use Filepath if provided, else current directory
		dir := fwo.Filepath
		if dir == "" {
			dir = "."
		}
		// Ensure directory exists
		if mkErr := os.MkdirAll(dir, 0755); mkErr != nil {
			return nil, fmt.Errorf("failed to create directory %s: %w", dir, mkErr)
		}
		filePath = filepath.Join(dir, fileName)
		L().Debug("creating regular file", String("filePath", filePath))
		if !fwo.OverwriteFile {
			if _, err = os.Stat(filePath); err == nil {
				return nil, fmt.Errorf("file already exists: %s", filePath)
			}
		}
		file, err = os.Create(filePath)
		if err != nil {
			return nil, fmt.Errorf("failed to create file: %w", err)
		}
	}

	defer func() {
		if closeErr := file.Close(); closeErr != nil {
			L().Warn("failed to Close file", String("filePath", filePath), Error(closeErr))
		}
	}()

	var writer io.Writer = file
	var gzipWriter *gzip.Writer

	// Add gzip compression if requested
	if fwo.UseGzip {
		L().Debug("enabling gzip compression for file", String("filePath", filePath))
		gzipWriter = gzip.NewWriter(file)
		defer func() {
			if closeErr := gzipWriter.Close(); closeErr != nil {
				L().Warn("failed to Close gzip writer", Error(closeErr))
			}
		}()
		writer = gzipWriter
	}

	L().Debug("writing data to file", String("filePath", filePath), String("fileName", fileName))

	// Write data using the provided write function
	err = writeFunc(writer)
	if err != nil {
		return nil, fmt.Errorf("failed to write data to %s: %w", filePath, err)
	}

	L().Info("file written successfully", String("filePath", filePath), String("fileName", fileName))

	return &FileWriteResult{
		Filepath: filePath,
		Filename: fileName,
	}, nil
}

// RemoveFile safely removes a file with improved error handling and logging
func (fwr FileWriteResult) RemoveFile() error {
	// Handle empty file path gracefully
	if fwr.Filepath == "" {
		L().Debug("no file path specified for removal, skipping")
		return nil // Nothing to remove
	}

	// Check if file exists before attempting removal
	if _, err := os.Stat(fwr.Filepath); os.IsNotExist(err) {
		L().Debug("file does not exist, skipping removal", String("filePath", fwr.Filepath))
		return nil // File doesn't exist, nothing to remove
	}

	// Attempt to remove the file
	if err := os.Remove(fwr.Filepath); err != nil {
		// Check if file was already removed (race condition)
		if os.IsNotExist(err) {
			L().Debug("file was already removed", String("filePath", fwr.Filepath))
			return nil
		}

		// Check for permission errors
		if os.IsPermission(err) {
			L().Error("insufficient permissions to remove file",
				String("filePath", fwr.Filepath),
				String("fileName", fwr.Filename),
				Error(err))
			return fmt.Errorf("insufficient permissions to remove file '%s': %w", fwr.Filename, err)
		}

		// Handle other file system errors
		L().Error("failed to remove file",
			String("filePath", fwr.Filepath),
			String("fileName", fwr.Filename),
			Error(err))
		return fmt.Errorf("failed to remove file '%s': %w", fwr.Filename, err)
	}

	// Log successful removal
	L().Info("file removed successfully",
		String("filePath", fwr.Filepath),
		String("fileName", fwr.Filename))

	return nil
}
