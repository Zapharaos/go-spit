# Makefile for go-spit

.PHONY: help test coverage lint fmt

help:
	@echo "Usage: make <target>"
	@echo "Targets:"
	@echo "  help      Show this help message"
	@echo "  test      Run tests (excluding examples)"
	@echo "  coverage  Run tests with coverage (excluding examples)"
	@echo "  lint      Run golangci-lint (excluding examples)"
	@echo "  fmt       Tries to automatically fix linting errors (excluding examples)"

# Run tests, excluding the examples package
test:
	go test $(shell go list ./... | grep -v '/gen' | grep -v '/test' | grep -v '/examples')

# Run tests with coverage, excluding the examples package
coverage:
	go test -cover $(shell go list ./... | grep -v '/gen' | grep -v '/test' | grep -v '/examples')

# Run golangci-lint, excluding the examples package
lint:
	golangci-lint run -exclude-dirs=examples

# Run with fix to automatically fix issues
fmt:
	golangci-lint run -exclude-dirs=examples --fix
