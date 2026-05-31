# Logging

go-spit emits structured log messages during export. By default it uses a built-in logger backed
by the standard library `log` package, but you can plug in your own logger and control verbosity.

## Log levels

`LogLevel` controls how much is logged. Higher levels include all lower levels:

| Level        | Description            |
|--------------|------------------------|
| `LevelOff`   | Logging disabled       |
| `LevelError` | Errors only            |
| `LevelWarn`  | Warnings and errors    |
| `LevelInfo`  | Informational (default) |
| `LevelDebug` | Everything, verbose    |

Manage the level with the package-level helpers:

```go
spit.SetLogLevel(spit.LevelDebug) // increase verbosity
level := spit.GetLogLevel()       // read the current level
spit.DisableLogger()              // equivalent to LevelOff
```

`HasLogLevel(level)` reports whether a given level is currently enabled.

## The default logger

Out of the box, `StdLogger` writes to **stdout** at `INFO` level using the standard library `log`
package. Messages are formatted as `[LEVEL] message | key=value …`.

To restore the defaults at any time:

```go
spit.ResetLogger() // reinstate StdLogger and LevelInfo
```

## Using your own logger

Implement the `Logger` interface to route go-spit logs into your application's logging stack
(Zap, Logrus, slog, …):

```go
type Logger interface {
	Debug(msg string, fields ...Field)
	Info(msg string, fields ...Field)
	Warn(msg string, fields ...Field)
	Error(msg string, fields ...Field)
}
```

Each message carries optional structured `Field` values. Construct fields with the provided
helpers: `String`, `Int`, `Bool`, `Error` and `Any`.

Install your logger with `SetLogger`, which returns a function that restores the previous logger:

```go
restore := spit.SetLogger(myLogger)
defer restore() // put the previous logger back when done
```

### Example adapter

A minimal adapter that forwards to the standard library:

```go
type myLogger struct{}

func (myLogger) Debug(msg string, fields ...spit.Field) { log.Println("DEBUG", msg, fields) }
func (myLogger) Info(msg string, fields ...spit.Field)  { log.Println("INFO", msg, fields) }
func (myLogger) Warn(msg string, fields ...spit.Field)  { log.Println("WARN", msg, fields) }
func (myLogger) Error(msg string, fields ...spit.Field) { log.Println("ERROR", msg, fields) }

func main() {
	restore := spit.SetLogger(myLogger{})
	defer restore()
	// ... perform exports ...
}
```
