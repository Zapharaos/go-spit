package spit

import (
	"os"
	"reflect"
	"testing"
)

func TestString(t *testing.T) {
	field := String("key", "value")
	if field.Key != "key" {
		t.Errorf("String() key = %q, want %q", field.Key, "key")
	}
	if field.Value != "value" {
		t.Errorf("String() value = %q, want %q", field.Value, "value")
	}
}

func TestError(t *testing.T) {
	err := os.ErrNotExist
	field := Error(err)
	if field.Key != "error" {
		t.Errorf("Error() key = %q, want %q", field.Key, "error")
	}
	if field.Value != err {
		t.Errorf("Error() value = %v, want %v", field.Value, err)
	}
}

func TestInt(t *testing.T) {
	field := Int("count", 42)
	if field.Key != "count" {
		t.Errorf("Int() key = %q, want %q", field.Key, "count")
	}
	if field.Value != 42 {
		t.Errorf("Int() value = %v, want %v", field.Value, 42)
	}
}

func TestBool(t *testing.T) {
	field := Bool("enabled", true)
	if field.Key != "enabled" {
		t.Errorf("Bool() key = %q, want %q", field.Key, "enabled")
	}
	if field.Value != true {
		t.Errorf("Bool() value = %v, want %v", field.Value, true)
	}
}

func TestAny(t *testing.T) {
	value := map[string]int{"test": 123}
	field := Any("data", value)
	if field.Key != "data" {
		t.Errorf("Any() key = %q, want %q", field.Key, "data")
	}
	if !reflect.DeepEqual(field.Value, value) {
		t.Errorf("Any() value = %v, want %v", field.Value, value)
	}
}

func TestL(t *testing.T) {
	// Set a test logger
	mockLogger := &MockLogger{}
	restore := SetLogger(mockLogger)
	defer restore()

	logger := L()
	if logger != mockLogger {
		t.Errorf("L() = %v, want %v", logger, mockLogger)
	}
}

func TestSetLogger(t *testing.T) {
	// Save original logger
	mockLogger := &MockLogger{}

	// Test setting logger and restore function
	restore := SetLogger(mockLogger)
	defer restore()
	if L() != mockLogger {
		t.Errorf("SetLogger() did not set logger correctly")
	}
}

func TestSetLogLevel(t *testing.T) {
	// Save original log level
	originalLevel := GetLogLevel()
	defer func() { SetLogLevel(originalLevel) }()

	SetLogLevel(LevelDebug)
	if GetLogLevel() != LevelDebug {
		t.Errorf("SetLogLevel(LevelDebug) = %v, want %v", GetLogLevel(), LevelDebug)
	}

	SetLogLevel(LevelError)
	if GetLogLevel() != LevelError {
		t.Errorf("SetLogLevel(LevelError) = %v, want %v", GetLogLevel(), LevelError)
	}
}

func TestHasLogLevel(t *testing.T) {
	// Save original log level
	originalLevel := GetLogLevel()
	defer func() { SetLogLevel(originalLevel) }()

	tests := []struct {
		name         string
		currentLevel LogLevel
		testLevel    LogLevel
		expected     bool
	}{
		{"Off level blocks all", LevelOff, LevelError, false},
		{"Error allows Error", LevelError, LevelError, true},
		{"Error blocks Warn", LevelError, LevelWarn, false},
		{"Info allows Error", LevelInfo, LevelError, true},
		{"Info allows Warn", LevelInfo, LevelWarn, true},
		{"Info allows Info", LevelInfo, LevelInfo, true},
		{"Info blocks Debug", LevelInfo, LevelDebug, false},
		{"Debug allows all", LevelDebug, LevelDebug, true},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			SetLogLevel(tt.currentLevel)
			result := HasLogLevel(tt.testLevel)
			if result != tt.expected {
				t.Errorf("HasLogLevel(%v) with level %v = %v, want %v",
					tt.testLevel, tt.currentLevel, result, tt.expected)
			}
		})
	}
}

func TestDisableLogger(t *testing.T) {
	// Save original log level
	originalLevel := GetLogLevel()
	defer func() { SetLogLevel(originalLevel) }()

	DisableLogger()
	if GetLogLevel() != LevelOff {
		t.Errorf("DisableLogger() did not set level to LevelOff, got %v", GetLogLevel())
	}
}

func TestResetLogger(t *testing.T) {
	// Save original state
	originalLevel := GetLogLevel()
	defer func() { SetLogLevel(originalLevel) }()

	// Change state
	mockLogger := &MockLogger{}
	restore := SetLogger(mockLogger)
	defer restore()
	SetLogLevel(LevelDebug)

	// Reset
	ResetLogger()

	if _, ok := L().(*StdLogger); !ok {
		t.Errorf("ResetLogger() did not reset logger to StdLogger, got %T", L())
	}
	if GetLogLevel() != LevelInfo {
		t.Errorf("ResetLogger() did not reset level to LevelInfo, got %v", GetLogLevel())
	}
}
