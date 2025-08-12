# Makefile for go-spit

.PHONY: help test coverage lint fmt

help:
	@echo "Usage: make <target>"
	@echo "Targets:"
	@echo "  help      Show this help message"
	@echo "  test      Run tests"
	@echo "  coverage  Run tests with coverage"
	@echo "  lint      Run golangci-lint"
	@echo "  fmt       Tries to automatically fix linting errors"

# Run tests, excluding the examples package
test:
	go test $(shell go list ./... | grep -v '/gen' | grep -v '/test')

# Run tests with coverage, excluding the examples package
coverage:
	go test -cover $(shell go list ./... | grep -v '/gen' | grep -v '/test')

# Run golangci-lint, excluding the examples package
lint:
	golangci-lint run

# Run with fix to automatically fix issues
fmt:
	golangci-lint run --fix
