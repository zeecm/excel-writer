.PHONY: install format check test

PACKAGE = "excel_writer"

install:
	pip install -e .[dev]
	pre-commit autoupdate
	pre-commit install

check:
	-pylint $(PACKAGE)
	pyright ./ tests/

test:
	pytest --cov=./ tests/

format:
	pycln .
	black . --preview
	isort .

make-acknowledgement-form:
	pyinstaller --add-data "acknowledgement_form/gui/resources;acknowledgement_form/gui/resources" --add-data "acknowledgement_form/template;acknowledgement_form/template" --name "acknowledgement-form-generator" --icon "acknowledgement_form\gui\resources\mencast_logo_ico.ico" ./acknowledgement_form/main.py
