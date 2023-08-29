.PHONY: install format check test

PACKAGE = "excel_writer"

install:
	pip install -e .[dev]
	pre-commit autoupdate
	pre-commit install

check:
	-pylint ./
	pyright ./ tests/

test:
	pytest --cov=./ tests/

format:
	pycln .
	black .
	isort .