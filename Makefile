.PHONY: install run test lint check

install:
	python -m pip install -e ".[dev]"

run:
	streamlit run streamlit_app.py

test:
	pytest

lint:
	ruff check .

check: lint test
	python -m compileall -q src streamlit_app.py
