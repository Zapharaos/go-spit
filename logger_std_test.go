package spit

import (
	"bytes"
	"log"
	"os"
	"strings"
	"testing"
)

func TestStdLogger_Debug(t *testing.T) {
	// Save original log level
	originalLevel := GetLogLevel()
	defer func() { SetLogLevel(originalLevel) }()

	// Capture log output
	var buf bytes.Buffer
	log.SetOutput(&buf)
	defer log.SetOutput(os.Stderr)

	logger := &StdLogger{}

	// Test with debug level enabled
	SetLogLevel(LevelDebug)
	logger.Debug("test message", String("key", "value"))

	output := buf.String()
	if !strings.Contains(output, "[DEBUG]") {
		t.Errorf("Debug() output missing [DEBUG] tag: %s", output)
	}
	if !strings.Contains(output, "test message") {
		t.Errorf("Debug() output missing message: %s", output)
	}
	if !strings.Contains(output, "key=value") {
		t.Errorf("Debug() output missing field: %s", output)
	}

	// Test with debug level disabled
	buf.Reset()
	SetLogLevel(LevelInfo)
	logger.Debug("should not appear")

	if buf.Len() > 0 {
		t.Errorf("Debug() should not log when level is higher: %s", buf.String())
	}
}

func TestStdLogger_Info(t *testing.T) {
	// Save original log level
	originalLevel := GetLogLevel()
	defer func() { SetLogLevel(originalLevel) }()

	// Capture log output
	var buf bytes.Buffer
	log.SetOutput(&buf)
	defer log.SetOutput(os.Stderr)

	logger := &StdLogger{}

	// Test with info level enabled
	SetLogLevel(LevelInfo)
	logger.Info("test info", Int("count", 5))

	output := buf.String()
	if !strings.Contains(output, "[INFO]") {
		t.Errorf("Info() output missing [INFO] tag: %s", output)
	}
	if !strings.Contains(output, "test info") {
		t.Errorf("Info() output missing message: %s", output)
	}
	if !strings.Contains(output, "count=5") {
		t.Errorf("Info() output missing field: %s", output)
	}

	// Test with info level disabled
	buf.Reset()
	SetLogLevel(LevelError)
	logger.Info("should not appear")

	if buf.Len() > 0 {
		t.Errorf("Info() should not log when level is higher: %s", buf.String())
	}
}

func TestStdLogger_Warn(t *testing.T) {
	// Save original log level
	originalLevel := GetLogLevel()
	defer func() { SetLogLevel(originalLevel) }()

	// Capture log output
	var buf bytes.Buffer
	log.SetOutput(&buf)
	defer log.SetOutput(os.Stderr)

	logger := &StdLogger{}

	// Test with warn level enabled
	SetLogLevel(LevelWarn)
	logger.Warn("test warning", Bool("urgent", true))

	output := buf.String()
	if !strings.Contains(output, "[WARN]") {
		t.Errorf("Warn() output missing [WARN] tag: %s", output)
	}
	if !strings.Contains(output, "test warning") {
		t.Errorf("Warn() output missing message: %s", output)
	}
	if !strings.Contains(output, "urgent=true") {
		t.Errorf("Warn() output missing field: %s", output)
	}

	// Test with warn level disabled
	buf.Reset()
	SetLogLevel(LevelError)
	logger.Warn("should not appear")

	if buf.Len() > 0 {
		t.Errorf("Warn() should not log when level is higher: %s", buf.String())
	}
}

func TestStdLogger_Error(t *testing.T) {
	// Save original log level
	originalLevel := GetLogLevel()
	defer func() { SetLogLevel(originalLevel) }()

	// Capture log output
	var buf bytes.Buffer
	log.SetOutput(&buf)
	defer log.SetOutput(os.Stderr)

	logger := &StdLogger{}

	// Test with error level enabled
	SetLogLevel(LevelError)
	logger.Error("test error", Error(os.ErrNotExist))

	output := buf.String()
	if !strings.Contains(output, "[ERROR]") {
		t.Errorf("Error() output missing [ERROR] tag: %s", output)
	}
	if !strings.Contains(output, "test error") {
		t.Errorf("Error() output missing message: %s", output)
	}
	if !strings.Contains(output, "error=") {
		t.Errorf("Error() output missing error field: %s", output)
	}

	// Test with error level disabled (LevelOff)
	buf.Reset()
	SetLogLevel(LevelOff)
	logger.Error("should not appear")

	if buf.Len() > 0 {
		t.Errorf("Error() should not log when level is LevelOff: %s", buf.String())
	}
}

func TestStdLogger_logWithFields(t *testing.T) {
	// Capture log output
	var buf bytes.Buffer
	log.SetOutput(&buf)
	defer log.SetOutput(os.Stderr)

	logger := &StdLogger{}

	// Test without fields
	logger.logWithFields("TEST", "message without fields")
	output := buf.String()
	if !strings.Contains(output, "[TEST] message without fields") {
		t.Errorf("logWithFields() without fields failed: %s", output)
	}

	// Test with multiple fields
	buf.Reset()
	logger.logWithFields("TEST", "message with fields",
		String("key1", "value1"),
		Int("key2", 42),
		Bool("key3", false))

	output = buf.String()
	if !strings.Contains(output, "[TEST] message with fields") {
		t.Errorf("logWithFields() missing message: %s", output)
	}
	if !strings.Contains(output, "key1=value1") {
		t.Errorf("logWithFields() missing first field: %s", output)
	}
	if !strings.Contains(output, "key2=42") {
		t.Errorf("logWithFields() missing second field: %s", output)
	}
	if !strings.Contains(output, "key3=false") {
		t.Errorf("logWithFields() missing third field: %s", output)
	}
}
