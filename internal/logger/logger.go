package logger

var (
	_logger Logger = &StdLogger{} // Default logger using standard log package
)

// L is used to access the global logger singleton
func L() Logger {
	return _logger
}

// ReplaceLogger replaces the global logger and returns a function to restore the previous one
func ReplaceLogger(logger Logger) func() {
	prev := _logger
	_logger = logger
	return func() { ReplaceLogger(prev) }
}

// Logger interface compatible with zap's signature
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
