# Trello Board CSV Exporter

This script exports data from a Trello board to a CSV file.

## Prerequisites

- **Trello API Key and Token**: Obtain these from the [Trello Developer Page](https://trello.com/power-ups/admin/) by creating/selecting a 'Power-Up'. Populate the `.env` file with your credentials based on `.env_example`.

## Virtual Environment (venv)

Set up a virtual environment and install dependencies:

```bash
python -m venv venv
source venv/bin/activate  # On Linux/macOS
.\venv\Scripts\activate   # On Windows
pip install -r requirements.txt
```

## Usage

Run the script and follow the prompts to select the board:

```bash
python -m trello_exporter
```

The CSV file will be generated in the `./csv` directory.
