BEHAVE = behave
MAKE   = make
TWINE  = twine
UV     = uv

.PHONY: accept build clean cleandocs coverage docs install lint opendocs sdist sync
.PHONY: test test-upload typecheck wheel

help:
	@echo "Please use \`make <target>' where <target> is one or more of"
	@echo "  accept       run acceptance tests using behave"
	@echo "  build        generate both sdist and wheel suitable for upload to PyPI"
	@echo "  clean        delete intermediate work product and start fresh"
	@echo "  cleandocs    delete intermediate documentation files"
	@echo "  coverage     run pytest with coverage"
	@echo "  docs         generate documentation"
	@echo "  lint         run Ruff"
	@echo "  opendocs     open browser to local version of documentation"
	@echo "  register     update metadata (README.rst) on PyPI"
	@echo "  sdist        generate a source distribution into dist/"
	@echo "  sync         create/update the UV-managed virtual environment"
	@echo "  test         run unit tests using pytest"
	@echo "  test-upload  upload distribution to TestPyPI"
	@echo "  typecheck    run ty"
	@echo "  upload       upload distribution tarball to PyPI"
	@echo "  wheel        generate a binary distribution into dist/"

accept:
	$(UV) run --group test $(BEHAVE) --stop

build:
	$(UV) build

clean:
	# find . -type f -name \*.pyc -exec rm {} \;
	fd -e pyc -I -x rm
	rm -rf dist *.egg-info .coverage .DS_Store

cleandocs:
	$(MAKE) -C docs clean

coverage:
	$(UV) run --group test pytest --cov-report term-missing --cov=docx tests/

docs:
	$(MAKE) -C docs html

install:
	$(MAKE) sync

lint:
	$(UV) run --group lint ruff check .

opendocs:
	open docs/.build/html/index.html

sdist:
	$(UV) build --sdist

sync:
	$(UV) sync

test:
	$(UV) run --group test pytest -x

test-upload: sdist wheel
	$(UV) run --group publish $(TWINE) upload --repository testpypi dist/*

typecheck:
	$(UV) run --group typing ty check

upload: clean sdist wheel
	$(UV) run --group publish $(TWINE) upload dist/*

wheel:
	$(UV) build --wheel
