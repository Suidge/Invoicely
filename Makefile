.PHONY: run build

run:
	PYTHONPATH=src python3 src/invoicely/invoice_sorter_native.py

build:
	./scripts/build_app.sh
