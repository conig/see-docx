PREFIX ?= $(HOME)/.local
PYTHON ?= python3

.PHONY: test check comments-smoke install uninstall

test:
	PYTHONPATH=src $(PYTHON) -m unittest discover -s tests -v

check:
	PYTHONPATH=src $(PYTHON) -m compileall -q src tests
	$(MAKE) test

comments-smoke:
	PYTHONPATH=src $(PYTHON) tests/ui_comments_smoke.py

install:
	install -d "$(PREFIX)/bin" "$(PREFIX)/lib/see-docx" "$(PREFIX)/share/applications"
	install -m 755 bin/see-docx "$(PREFIX)/bin/see-docx"
	cp -R src/see_docx "$(PREFIX)/lib/see-docx/"
	install -m 644 data/io.github.conig.seedocx.desktop "$(PREFIX)/share/applications/"

uninstall:
	rm -f "$(PREFIX)/bin/see-docx" "$(PREFIX)/share/applications/io.github.conig.seedocx.desktop"
	rm -rf "$(PREFIX)/lib/see-docx"
