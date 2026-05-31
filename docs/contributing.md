# Contributing

We appreciate your interest in contributing to go-spit. Community contributions — bug fixes, new
features, documentation improvements — mean a lot to us. This page summarizes how to get started;
the authoritative source is
[`CONTRIBUTING.md`](https://github.com/Zapharaos/go-spit/blob/main/CONTRIBUTING.md) in the
repository.

## Getting started

1. Fork the repository and clone your fork.
2. Make your changes in a feature branch.
3. Add or update tests for your change.
4. Make sure the test suite and linters pass.
5. Open a pull request describing your change.

Not sure where to begin? Look through the
[help-wanted issues](https://github.com/Zapharaos/go-spit/issues?q=is%3Aissue%20state%3Aopen%20label%3A%22help%20wanted%22),
or improve documentation and tutorials.

## Development environment

The project uses a `Makefile` for common tasks. Install the development dependencies first:

```sh
make dev-deps
```

| Command          | Description                                          |
|------------------|------------------------------------------------------|
| `make dev-deps`  | Install development dependencies (linter, mockgen).  |
| `make test-unit` | Run unit tests and generate a coverage report.       |
| `make lint`      | Run `golangci-lint`.                                  |
| `make fmt`       | Automatically fix lint issues where possible.        |
| `make mocks`     | Regenerate interface mocks (`go generate ./...`).    |

### Running tests

```sh
make test-unit
```

If you get errors, make sure you are using a supported version of Go
(**1.24.1** or newer).

### Linting

```sh
make lint
```

`make fmt` can fix many issues automatically.

### Mocks

Interface mocks (for example `logger_mock.go` and `spreadsheet_mock.go`) are generated with
`mockgen`. After changing a mocked interface, regenerate them:

```sh
make mocks
```

## Working on the documentation

This site is built with [MkDocs](https://www.mkdocs.org/) and the
[Material theme](https://squidfunk.github.io/mkdocs-material/). The Markdown sources live in the
[`docs/`](https://github.com/Zapharaos/go-spit/tree/main/docs) directory and the configuration is
in [`mkdocs.yml`](https://github.com/Zapharaos/go-spit/blob/main/mkdocs.yml).

Install the documentation dependencies:

```sh
pip install -r docs/requirements.txt
```

Preview the site locally with live reload:

```sh
mkdocs serve
```

Then open <http://127.0.0.1:8000/>. Build a static copy into the `site/` directory with:

```sh
mkdocs build
```

The documentation is **deployed automatically** to GitHub Pages by the
[`docs` workflow](https://github.com/Zapharaos/go-spit/blob/main/.github/workflows/docs.yml)
whenever changes land on the `main` branch.

## Reporting a bug

When filing an issue, please answer:

1. What version of go-spit are you using?
2. What did you do?
3. What did you expect to see?
4. What did you see instead?

## Suggesting a feature

Check the [issue list](https://github.com/Zapharaos/go-spit/issues) first to avoid duplicates. If
nothing matches, open a new issue describing the feature and how it should work.

## Code review process

The core team regularly reviews pull requests and provides feedback as soon as possible. After
receiving feedback, please respond within two weeks; otherwise the PR may be closed for inactivity.
