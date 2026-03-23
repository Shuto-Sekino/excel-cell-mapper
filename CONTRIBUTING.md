# Contributing to excel-cell-mapper

Thank you for your interest in contributing!

## Getting Started

### Prerequisites

- Python 3.13+
- [uv](https://github.com/astral-sh/uv)

### Setup

```bash
git clone https://github.com/shutosekino/excel-cell-mapper.git
cd excel-cell-mapper
uv sync --extra dev
bash scripts/initial-setup.sh
```

> **Note:** `scripts/initial-setup.sh` installs the git hooks (pre-push) required for this project. You must run it before starting development. It ensures lint checks run automatically before each push.

## Development Workflow

A `Makefile` is provided for common tasks:

| Command | Description |
|---------|-------------|
| `make test` | Run the test suite |
| `make lint` | Check for lint errors |
| `make format` | Apply code formatting |
| `make check` | Run lint, format check, and tests in one go |

### Running Tests

```bash
make test
```

With coverage:

```bash
uv run pytest --cov=src/excel_cell_mapper
```

### Linting and Formatting

This project uses [ruff](https://docs.astral.sh/ruff/) for linting and formatting.

```bash
make lint      # Check for lint errors
make format    # Apply formatting

# Or run everything at once:
make check
```

For auto-fixing lint errors:

```bash
uv run ruff check --fix src/ tests/
```

## Submitting a Pull Request

1. Fork the repository and create a branch from `main`.
2. Make your changes, adding tests for any new behavior.
3. Run `make check` to ensure all tests pass and there are no lint/format errors.
4. Open a pull request against `main`.

CI will automatically run tests and linting on your PR.

## Project Structure

```
src/excel_cell_mapper/   # Library source code
tests/                   # Test suite
.github/workflows/       # CI workflows (test, lint, publish)
```

## Reporting Issues

Please open an issue at https://github.com/shutosekino/excel-cell-mapper/issues.

## License

By contributing, you agree that your contributions will be licensed under the [MIT License](LICENSE).
