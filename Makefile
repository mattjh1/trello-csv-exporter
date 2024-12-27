.PHONY: install uninstall clean

# Install the package
install:
	python3 -m pip install --upgrade pip setuptools build
	python3 -m build
	python3 -m pip install dist/*.whl

# Uninstall the package
uninstall:
	python3 -m pip uninstall -y trello_exporter

# Clean build artifacts
clean:
	rm -rf build dist trello_exporter.egg-info
