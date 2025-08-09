package go_spit

import "github.com/Zapharaos/go-spit/internal/logger"

// SetLogger replaces the global logger and returns a function to restore the previous one
func SetLogger(newLogger logger.Logger) func() {
	return logger.ReplaceLogger(newLogger)
}
