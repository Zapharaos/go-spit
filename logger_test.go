package spit_test

import (
	"os"
	"reflect"
	"testing"

	"github.com/Zapharaos/go-spit"
	"github.com/Zapharaos/go-spit/mocks"
)

func TestString(t *testing.T) {
	field := spit.String("key", "value")
	if field.Key != "key" {
		t.Errorf("String() key = %q, want %q", field.Key, "key")
	}
	if field.Value != "value" {
		t.Errorf("String() value = %q, want %q", field.Value, "value")
	}
}

func TestError(t *testing.T) {
	err := os.ErrNotExist
	field := spit.Error(err)
	if field.Key != "error" {
		t.Errorf("Error() key = %q, want %q", field.Key, "error")
	}
	if field.Value != err {
		t.Errorf("Error() value = %v, want %v", field.Value, err)
	}
}

func TestInt(t *testing.T) {
	field := spit.Int("count", 42)
	if field.Key != "count" {
		t.Errorf("Int() key = %q, want %q", field.Key, "count")
	}
	if field.Value != 42 {
		t.Errorf("Int() value = %v, want %v", field.Value, 42)
	}
}

func TestBool(t *testing.T) {
	field := spit.Bool("enabled", true)
	if field.Key != "enabled" {
		t.Errorf("Bool() key = %q, want %q", field.Key, "enabled")
	}
	if field.Value != true {
		t.Errorf("Bool() value = %v, want %v", field.Value, true)
	}
}

func TestAny(t *testing.T) {
	value := map[string]int{"test": 123}
	field := spit.Any("data", value)
	if field.Key != "data" {
		t.Errorf("Any() key = %q, want %q", field.Key, "data")
	}
	if !reflect.DeepEqual(field.Value, value) {
		t.Errorf("Any() value = %v, want %v", field.Value, value)
	}
}

func TestL(t *testing.T) {
	// Set a test logger
	mockLogger := &mocks.MockLogger{}
	restore := spit.SetLogger(mockLogger)
	defer restore()

	logger := spit.L()
	if logger != mockLogger {
		t.Errorf("L() = %v, want %v", logger, mockLogger)
	}
}

func TestSetLogger(t *testing.T) {
	// Save original logger
	mockLogger := &mocks.MockLogger{}

	// Test setting logger and restore function
	restore := spit.SetLogger(mockLogger)
	defer restore()
	if spit.L() != mockLogger {
		t.Errorf("SetLogger() did not set logger correctly")
	}
}

func TestSetLogLevel(t *testing.T) {
	// Save original log level
	originalLevel := spit.GetLogLevel()
	defer func() { spit.SetLogLevel(originalLevel) }()

	spit.SetLogLevel(spit.LevelDebug)
	if spit.GetLogLevel() != spit.LevelDebug {
		t.Errorf("SetLogLevel(LevelDebug) = %v, want %v", spit.GetLogLevel(), spit.LevelDebug)
	}

	spit.SetLogLevel(spit.LevelError)
	if spit.GetLogLevel() != spit.LevelError {
		t.Errorf("SetLogLevel(LevelError) = %v, want %v", spit.GetLogLevel(), spit.LevelError)
	}
}

func TestHasLogLevel(t *testing.T) {
	// Save original log level
	originalLevel := spit.GetLogLevel()
	defer func() { spit.SetLogLevel(originalLevel) }()

	tests := []struct {
		name         string
		currentLevel spit.LogLevel
		testLevel    spit.LogLevel
		expected     bool
	}{
		{"Off level blocks all", spit.LevelOff, spit.LevelError, false},
		{"Error allows Error", spit.LevelError, spit.LevelError, true},
		{"Error blocks Warn", spit.LevelError, spit.LevelWarn, false},
		{"Info allows Error", spit.LevelInfo, spit.LevelError, true},
		{"Info allows Warn", spit.LevelInfo, spit.LevelWarn, true},
		{"Info allows Info", spit.LevelInfo, spit.LevelInfo, true},
		{"Info blocks Debug", spit.LevelInfo, spit.LevelDebug, false},
		{"Debug allows all", spit.LevelDebug, spit.LevelDebug, true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			spit.SetLogLevel(tt.currentLevel)
			result := spit.HasLogLevel(tt.testLevel)
			if result != tt.expected {
				t.Errorf("HasLogLevel(%v) with level %v = %v, want %v",
					tt.testLevel, tt.currentLevel, result, tt.expected)
			}
		})
	}
}

func TestDisableLogger(t *testing.T) {
	// Save original log level
	originalLevel := spit.GetLogLevel()
	defer func() { spit.SetLogLevel(originalLevel) }()

	spit.DisableLogger()
	if spit.GetLogLevel() != spit.LevelOff {
		t.Errorf("DisableLogger() did not set level to LevelOff, got %v", spit.GetLogLevel())
	}
}

func TestResetLogger(t *testing.T) {
	// Save original state
	originalLevel := spit.GetLogLevel()
	defer func() { spit.SetLogLevel(originalLevel) }()

	// Change state
	mockLogger := &mocks.MockLogger{}
	restore := spit.SetLogger(mockLogger)
	defer restore()
	spit.SetLogLevel(spit.LevelDebug)

	// Reset
	spit.ResetLogger()

	if _, ok := spit.L().(*spit.StdLogger); !ok {
		t.Errorf("ResetLogger() did not reset logger to StdLogger, got %T", spit.L())
	}
	if spit.GetLogLevel() != spit.LevelInfo {
		t.Errorf("ResetLogger() did not reset level to LevelInfo, got %v", spit.GetLogLevel())
	}
}
