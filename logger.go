//go:generate mockgen -destination=logger_mock.go -package=spit . Logger

package spit

// LogLevel represents the severity level for logging. Used to filter log output.
type LogLevel int

const (
	LevelOff   LogLevel = iota // Logging disabled
	LevelError                 // Error level logs
	LevelWarn                  // Warning level logs
	LevelInfo                  // Informational logs
	LevelDebug                 // Debug level logs
)

var (
	_logger   Logger   = &StdLogger{} // Global logger instance
	_logLevel LogLevel = LevelInfo    // Default log level
)

// Logger defines the interface for logging implementations.
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

// String returns a Field with a string value.
func String(key, val string) Field {
	return Field{Key: key, Value: val}
}

// Error returns a Field for an error value.
func Error(err error) Field {
	return Field{Key: "error", Value: err}
}

// Int returns a Field with an int value.
func Int(key string, val int) Field {
	return Field{Key: key, Value: val}
}

// Bool returns a Field with a bool value.
func Bool(key string, val bool) Field {
	return Field{Key: key, Value: val}
}

// Any returns a Field with any value type.
func Any(key string, val interface{}) Field {
	return Field{Key: key, Value: val}
}

// L returns the global logger instance.
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
