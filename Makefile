.PHONY: install uninstall clean build sign upload release

# Install the package
install:
	python3 -m pip install --upgrade pip setuptools build twine
	python3 -m build
	python3 -m pip install dist/*.whl

# Uninstall the package
uninstall:
	python3 -m pip uninstall -y trello_csv

# Clean build artifacts
clean:
	rm -rf build dist trello_csv.egg-info

# Build the package
build:
	python3 -m build

# Sign the package files with GPG
sign:
	@echo "Signing package files..."
	@for file in dist/*.{tar.gz,whl}; do \
		if [ -f "$$file" ]; then \
			gpg --detach-sign -a $$file; \
			echo "Signed $$file"; \
		fi; \
	done

# Upload the package and signatures to PyPI
upload:
	twine upload dist/*

# Full release process: clean, build, sign, and upload
release: clean build sign upload
