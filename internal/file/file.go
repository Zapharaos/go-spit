package file

import (
	"compress/gzip"
	"fmt"
	"io"
	"os"
	"strings"
	"unicode"

	"github.com/Zapharaos/go-spit/internal/logger"
)

// WriteOptions contains generic options for file writing
type WriteOptions struct {
	Filename      string // Desired filename (without extension)
	Extension     string // File extension (e.g., ".csv", ".json")
	UseGzip       bool   // Optional: compress with gzip
	OverwriteFile bool   // Optional: overwrite existing file (default: false)
}

// WriteResult contains the result of file writing operation
type WriteResult struct {
	FilePath string // Full path to the created file
	FileName string // Final filename (including any modifications)
}

// SanitizeFilename sanitizes a string to be safe for use as a filename.
func (opt WriteOptions) SanitizeFilename() string {
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
	}, opt.Filename)

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

// WriteToFile writes data to a file with generic options and returns file info
func (opt WriteOptions) WriteToFile(writeFunc func(io.Writer) error) (*WriteResult, error) {

	// Sanitize the filename to ensure it's safe for use
	opt.Filename = strings.ToLower(opt.SanitizeFilename())

	// Construct filename with extension
	extension := "." + strings.ToLower(opt.Extension)
	fileName := opt.Filename + extension
	tempFilePattern := opt.Filename + "_*" + extension

	if opt.UseGzip {
		fileName += ".gz"
		tempFilePattern += ".gz"
	}

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
	if !opt.OverwriteFile {
		if _, err = os.Stat(tempPath); err == nil {
			if removeErr := os.Remove(tempPath); removeErr != nil {
				logger.L().Warn("failed to remove temp file during cleanup", logger.String("filePath", tempPath), logger.Error(removeErr))
			}
			return nil, fmt.Errorf("temp file already exists: %s", tempPath)
		}
	}

	var writer io.Writer = tempFile
	var gzipWriter *gzip.Writer

	// Add gzip compression if requested
	if opt.UseGzip {
		gzipWriter = gzip.NewWriter(tempFile)
		defer func() {
			if closeErr := gzipWriter.Close(); closeErr != nil {
				logger.L().Warn("failed to close gzip writer", logger.Error(closeErr))
			}
		}()
		writer = gzipWriter
	}

	// Write data using the provided write function
	err = writeFunc(writer)
	if err != nil {
		return nil, fmt.Errorf("failed to write data: %w", err)
	}

	return &WriteResult{
		FilePath: tempPath,
		FileName: fileName,
	}, nil
}

// RemoveFile safely removes a file
func (fwr WriteResult) RemoveFile() error {
	if fwr.FilePath == "" {
		return nil // Nothing to remove
	}

	if err := os.Remove(fwr.FilePath); err != nil && !os.IsNotExist(err) {
		logger.L().Warn("failed to remove export file", logger.String("filePath", fwr.FilePath), logger.Error(err))
		return fmt.Errorf("failed to remove export file: %w", err)
	}

	return nil
}
