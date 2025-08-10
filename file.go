// file.go - Generic file writing and management.
//
// This file provides utilities for writing data to files with options and managing it.

package go_spit

import (
	"compress/gzip"
	"fmt"
	"io"
	"os"
	"strings"
	"unicode"

	"github.com/Zapharaos/go-spit/internal/logger"
)

// FileWriteOptions contains generic options for file writing
type FileWriteOptions struct {
	Filename      string // Desired filename (without extension)
	UseGzip       bool   // Optional: compress with gzip
	OverwriteFile bool   // Optional: overwrite existing file (default: false)
	extension     string // Optional: File extension (e.g., ".csv", ".json")
}

// FileWriteResult contains the result of file writing operation
type FileWriteResult struct {
	FilePath string // Full path to the created file
	FileName string // Final filename (including any modifications)
}

// SanitizeFilename sanitizes a string to be safe for use as a filename.
func (fwo FileWriteOptions) SanitizeFilename() string {
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
func (fwo FileWriteOptions) writeToFile(writeFunc func(io.Writer) error) (*FileWriteResult, error) {
	// Sanitize the filename to ensure it's safe for use
	fwo.Filename = strings.ToLower(fwo.SanitizeFilename())

	if fwo.Filename == "" {
		return nil, fmt.Errorf("filename is empty after sanitization")
	}

	// Construct filename with extension
	extension := "." + strings.ToLower(fwo.extension)
	fileName := fwo.Filename + extension
	tempFilePattern := fwo.Filename + "_*" + extension

	if fwo.UseGzip {
		fileName += ".gz"
		tempFilePattern += ".gz"
	}

	logger.L().Debug("creating temp file", logger.String("pattern", tempFilePattern))

	// Create temporary file
	tempFile, err := os.CreateTemp("", tempFilePattern)
	if err != nil {
		return nil, fmt.Errorf("failed to create temp file: %w", err)
	}
	tempPath := tempFile.Name()

	defer func() {
		if closeErr := tempFile.Close(); closeErr != nil {
			logger.L().Warn("failed to close temp file", logger.String("filePath", tempPath), logger.Error(closeErr))
		}
	}()

	// Check if file already exists when we don't want to overwrite (shouldn't happen with temp files)
	if !fwo.OverwriteFile {
		if _, err = os.Stat(tempPath); err == nil {
			return nil, fmt.Errorf("temp file already exists: %s", tempPath)
		}
	}

	var writer io.Writer = tempFile
	var gzipWriter *gzip.Writer

	// Add gzip compression if requested
	if fwo.UseGzip {
		logger.L().Debug("enabling gzip compression for file", logger.String("filePath", tempPath))
		gzipWriter = gzip.NewWriter(tempFile)
		defer func() {
			if closeErr := gzipWriter.Close(); closeErr != nil {
				logger.L().Warn("failed to close gzip writer", logger.Error(closeErr))
			}
		}()
		writer = gzipWriter
	}

	logger.L().Debug("writing data to file", logger.String("filePath", tempPath), logger.String("fileName", fileName))

	// Write data using the provided write function
	err = writeFunc(writer)
	if err != nil {
		return nil, fmt.Errorf("failed to write data to %s: %w", tempPath, err)
	}

	logger.L().Info("file written successfully", logger.String("filePath", tempPath), logger.String("fileName", fileName))

	return &FileWriteResult{
		FilePath: tempPath,
		FileName: fileName,
	}, nil
}

// RemoveFile safely removes a file with improved error handling and logging
func (fwr FileWriteResult) RemoveFile() error {
	// Handle empty file path gracefully
	if fwr.FilePath == "" {
		logger.L().Debug("no file path specified for removal, skipping")
		return nil // Nothing to remove
	}

	// Check if file exists before attempting removal
	if _, err := os.Stat(fwr.FilePath); os.IsNotExist(err) {
		logger.L().Debug("file does not exist, skipping removal", logger.String("filePath", fwr.FilePath))
		return nil // File doesn't exist, nothing to remove
	}

	// Attempt to remove the file
	if err := os.Remove(fwr.FilePath); err != nil {
		// Check if file was already removed (race condition)
		if os.IsNotExist(err) {
			logger.L().Debug("file was already removed", logger.String("filePath", fwr.FilePath))
			return nil
		}

		// Check for permission errors
		if os.IsPermission(err) {
			logger.L().Error("insufficient permissions to remove file",
				logger.String("filePath", fwr.FilePath),
				logger.String("fileName", fwr.FileName),
				logger.Error(err))
			return fmt.Errorf("insufficient permissions to remove file '%s': %w", fwr.FileName, err)
		}

		// Handle other file system errors
		logger.L().Error("failed to remove file",
			logger.String("filePath", fwr.FilePath),
			logger.String("fileName", fwr.FileName),
			logger.Error(err))
		return fmt.Errorf("failed to remove file '%s': %w", fwr.FileName, err)
	}

	// Log successful removal
	logger.L().Info("file removed successfully",
		logger.String("filePath", fwr.FilePath),
		logger.String("fileName", fwr.FileName))

	return nil
}
