PREFIX ?= $(HOME)/.local
PYTHON ?= python3
GUI_TEST_RUNNER = scripts/run-headless-gui-test

.PHONY: test check comments-smoke rich-selection-smoke gui-smoke test-battery install uninstall

test:
	PYTHONPATH=src $(PYTHON) -m unittest discover -s tests -v

check:
	PYTHONPATH=src $(PYTHON) -m compileall -q src tests
	$(MAKE) test

comments-smoke:
	PYTHONPATH=src $(GUI_TEST_RUNNER) $(PYTHON) tests/ui_comments_smoke.py

rich-selection-smoke:
	PYTHONPATH=src $(GUI_TEST_RUNNER) sh -c '$(PYTHON) tests/ui_rich_selection_smoke.py && $(PYTHON) tests/ui_multi_page_table_copy_smoke.py && $(PYTHON) tests/ui_table_mapping_smoke.py && $(PYTHON) tests/ui_table_boundary_smoke.py'

gui-smoke:
	PYTHONPATH=src $(GUI_TEST_RUNNER) sh -c '$(PYTHON) tests/ui_comments_smoke.py && $(PYTHON) tests/ui_rich_selection_smoke.py && $(PYTHON) tests/ui_multi_page_table_copy_smoke.py && $(PYTHON) tests/ui_table_mapping_smoke.py && $(PYTHON) tests/ui_table_boundary_smoke.py'

test-battery: check
	$(MAKE) gui-smoke

install:
	install -d "$(PREFIX)/bin" "$(PREFIX)/lib/see-docx" "$(PREFIX)/share/applications"
	install -m 755 bin/see-docx "$(PREFIX)/bin/see-docx"
	cp -R src/see_docx "$(PREFIX)/lib/see-docx/"
	install -m 644 data/io.github.conig.seedocx.desktop "$(PREFIX)/share/applications/"

uninstall:
	rm -f "$(PREFIX)/bin/see-docx" "$(PREFIX)/share/applications/io.github.conig.seedocx.desktop"
	rm -rf "$(PREFIX)/lib/see-docx"
