package go_spit

import (
	"fmt"
	"log"
)

// LogLevel represents the severity level for logging
type LogLevel int

const (
	LevelOff LogLevel = iota
	LevelError
	LevelWarn
	LevelInfo
	LevelDebug
)

var (
	_logger   Logger   = &StdLogger{}
	_logLevel LogLevel = LevelInfo
)

// Logger interface that can be implemented by any logging library
// Compatible with popular loggers like Zap, Logrus, etc.
type Logger interface {
	Debug(msg string, fields ...Field)
	Info(msg string, fields ...Field)
	Warn(msg string, fields ...Field)
	Error(msg string, fields ...Field)
}

// Field represents a key-value pair for structured logging
type Field struct {
	Key   string
	Value interface{}
}

func String(key, val string) Field {
	return Field{Key: key, Value: val}
}

func Error(err error) Field {
	return Field{Key: "error", Value: err}
}

func Int(key string, val int) Field {
	return Field{Key: key, Value: val}
}

func Bool(key string, val bool) Field {
	return Field{Key: key, Value: val}
}

func Any(key string, val interface{}) Field {
	return Field{Key: key, Value: val}
}

// L singleton logger access
func L() Logger {
	return _logger
}

// SetLogger replaces the global logger and returns a function to restore the previous one
func SetLogger(newLogger Logger) func() {
	prev := _logger
	_logger = newLogger
	return func() { _logger = prev }
}

// SetLogLevel sets the global log level for logging
func SetLogLevel(level LogLevel) {
	_logLevel = level
}

// GetLogLevel returns the current log level
func GetLogLevel() LogLevel {
	return _logLevel
}

// HasLogLevel checks if the current log level allows the specified level
func HasLogLevel(level LogLevel) bool {
	return _logLevel != LevelOff && _logLevel >= level
}

// DisableLogger disables all logging output
func DisableLogger() {
	_logLevel = LevelOff
}

// ResetLogger resets the logger and log level to their defaults
func ResetLogger() {
	_logger = &StdLogger{}
	_logLevel = LevelInfo
}

// StdLogger implements Logger using the standard log package
type StdLogger struct{}

func (l *StdLogger) Debug(msg string, fields ...Field) {
	if !HasLogLevel(LevelDebug) {
		return
	}
	l.logWithFields("DEBUG", msg, fields...)
}

func (l *StdLogger) Info(msg string, fields ...Field) {
	if !HasLogLevel(LevelInfo) {
		return
	}
	l.logWithFields("INFO", msg, fields...)
}

func (l *StdLogger) Warn(msg string, fields ...Field) {
	if !HasLogLevel(LevelWarn) {
		return
	}
	l.logWithFields("WARN", msg, fields...)
}

func (l *StdLogger) Error(msg string, fields ...Field) {
	if !HasLogLevel(LevelError) {
		return
	}
	l.logWithFields("ERROR", msg, fields...)
}

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
