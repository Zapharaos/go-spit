# Makefile for go-spit

.PHONY: help mocks test coverage lint fmt dev-deps

help:
	@echo "Usage: make <target>"
	@echo "Targets:"
	@echo "  help      Show this help message"
	@echo "  dev-deps  Install development dependencies"
	@echo "  mocks     Generate mocks for interfaces"
	@echo "  test      Run tests"
	@echo "  coverage  Run tests with coverage"
	@echo "  lint      Run golangci-lint"
	@echo "  fmt       Tries to automatically fix linting errors"

# Install development dependencies
dev-deps:
	go install github.com/golangci/golangci-lint/v2/cmd/golangci-lint@v2.4.0
	go install go.uber.org/mock/mockgen@latest

# Generate mocks for interfaces in the mocks package
mocks:
	go generate ./...

# Run tests, excluding the mocks directory
test:
	go test $(shell go list ./... | grep -v '/mocks')

# Run tests with coverage, excluding the mocks directory
coverage:
	go test -cover $(shell go list ./... | grep -v '/mocks')

# Run golangci-lint, excluding the mocks directory
lint:
	golangci-lint run --skip-dirs mocks

# Run with fix to automatically fix issues
fmt:
	golangci-lint run --fix --skip-dirs mocks