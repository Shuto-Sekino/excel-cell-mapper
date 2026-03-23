.PHONY: lint format test check

lint:
	uv run ruff check src/ tests/

format:
	uv run ruff format src/ tests/

test:
	uv run pytest

check:
	uv run ruff check src/ tests/
	uv run ruff format --check src/ tests/
	uv run pytest
