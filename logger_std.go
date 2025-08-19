package spit

import (
	"fmt"
	"log"
)

// StdLogger implements Logger using the standard log package.
// It outputs log messages to the default logger, respecting the configured log level.
type StdLogger struct{}

// Debug logs a debug-level message with optional fields.
func (l *StdLogger) Debug(msg string, fields ...Field) {
	if !HasLogLevel(LevelDebug) {
		return
	}
	l.logWithFields("DEBUG", msg, fields...)
}

// Info logs an info-level message with optional fields.
func (l *StdLogger) Info(msg string, fields ...Field) {
	if !HasLogLevel(LevelInfo) {
		return
	}
	l.logWithFields("INFO", msg, fields...)
}

// Warn logs a warning-level message with optional fields.
func (l *StdLogger) Warn(msg string, fields ...Field) {
	if !HasLogLevel(LevelWarn) {
		return
	}
	l.logWithFields("WARN", msg, fields...)
}

// Error logs an error-level message with optional fields.
func (l *StdLogger) Error(msg string, fields ...Field) {
	if !HasLogLevel(LevelError) {
		return
	}
	l.logWithFields("ERROR", msg, fields...)
}

// logWithFields formats and outputs a log message with the given level, message, and fields.
// Fields are formatted as key=value pairs and appended to the log output.
func (l *StdLogger) logWithFields(level, msg string, fields ...Field) {
	if len(fields) == 0 {
		log.Printf("[%s] %s", level, msg)
		return
	}

	fieldStrs := make([]string, len(fields))
	for i, field := range fields {
		fieldStrs[i] = fmt.Sprintf("%s=%v", field.Key, field.Value)
	}

	fieldStr := ""
	for i, fs := range fieldStrs {
		if i > 0 {
			fieldStr += " "
		}
		fieldStr += fs
	}

	log.Printf("[%s] %s | %s", level, msg, fieldStr)
}
