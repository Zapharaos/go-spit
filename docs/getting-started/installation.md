# Installation

## Requirements

- **Go 1.24.1** or newer.
- go-spit uses [Go Modules](https://go.dev/wiki/Modules) to manage dependencies.

## Add go-spit to your project

From the root of your Go module, run:

```sh
go get github.com/Zapharaos/go-spit
```

This adds go-spit to your `go.mod` file and downloads it into the module cache.

## Import the package

The module path is `github.com/Zapharaos/go-spit`, and the package name is `spit`:

```go
import "github.com/Zapharaos/go-spit"
```

You then reference exported symbols through the `spit` package, for example
`spit.NewTable`, `spit.ExportCSV` or `spit.ExportXLSX`.

## Verify the installation

Create a small program and run it to confirm everything is wired up correctly:

```go
package main

import (
	"log"

	"github.com/Zapharaos/go-spit"
)

func main() {
	table := spit.NewTable(
		spit.DataSlice{{"hello": "world"}},
		spit.Columns{spit.NewColumn("hello", "Hello")},
		true,
	)

	result, err := spit.ExportCSV(",", table, spit.FileWriteParams{Filename: "hello"})
	if err != nil {
		log.Fatal(err)
	}
	defer result.RemoveFile()

	log.Printf("created %s", result.Filepath)
}
```

```sh
go run .
```

If a `hello.csv` file is created, you are ready to continue with the
[Quick Start](quickstart.md).
