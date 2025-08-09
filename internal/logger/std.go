package logger

import (
	"fmt"
	"log"
)

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
