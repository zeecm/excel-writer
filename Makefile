.PHONY: install format check test

PACKAGE = "excel_writer"

install:
	pip install -e .[dev]
	pre-commit autoupdate
	pre-commit install

check:
	-pylint $(PACKAGE)
	pyright $(PACKAGE) tests/

test:
	pytest --cov=$(PACKAGE) tests/

format:
	pycln .
	black .
	isort .