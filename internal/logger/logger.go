package logger

import (
	"fmt"
	"log"
)

var (
	_globalLogger Logger = &StdLogger{} // Default logger using standard log package
)

// Field represents a key-value pair for structured logging
type Field struct {
	Key   string
	Value interface{}
}

// Logger interface compatible with zap's signature
type Logger interface {
	Debug(msg string, fields ...Field)
	Info(msg string, fields ...Field)
	Warn(msg string, fields ...Field)
	Error(msg string, fields ...Field)
}

// L is used to access the global logger singleton
func L() Logger {
	return _globalLogger
}

// ReplaceGlobals replaces the global logger and returns a function to restore the previous one
func ReplaceGlobals(logger Logger) func() {
	prev := _globalLogger
	_globalLogger = logger
	return func() { ReplaceGlobals(prev) }
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

// StdLogger implements Logger using the standard log package
type StdLogger struct{}

func (l *StdLogger) Debug(msg string, fields ...Field) {
	l.logWithFields("DEBUG", msg, fields...)
}

func (l *StdLogger) Info(msg string, fields ...Field) {
	l.logWithFields("INFO", msg, fields...)
}

func (l *StdLogger) Warn(msg string, fields ...Field) {
	l.logWithFields("WARN", msg, fields...)
}

func (l *StdLogger) Error(msg string, fields ...Field) {
	l.logWithFields("ERROR", msg, fields...)
}

func (l *StdLogger) logWithFields(level, msg string, fields ...Field) {
	if len(fields) == 0 {
		log.Printf("%s: %s", level, msg)
		return
	}

	fieldStrs := make([]string, len(fields))
	for i, field := range fields {
		fieldStrs[i] = fmt.Sprintf("%s=%v", field.Key, field.Value)
	}
	log.Printf("%s: %s [%s]", level, msg, fmt.Sprintf("%s", fieldStrs))
}
