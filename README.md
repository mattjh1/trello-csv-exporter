# Trello CSV Exporter

A simple command-line tool to export Trello board data into a CSV-compatible Excel file. You can choose to save the file locally or upload it directly to an AWS S3 bucket.

## Features

- **Export Trello data**: Export your Trello board’s cards, lists, and labels into an Excel sheet.
- **Save locally or upload to S3**: Save the generated file to a local directory or upload it directly to AWS S3.
- **AWS profile support**: Optionally configure AWS profiles for seamless S3 uploads.

## Installation

### 1. Install the tool via pip

You can install the tool from PyPI:

```bash
pip install trello-csv
```

### 2. Set up your credentials

- **Trello API Key and Token**: You will need your Trello API key and token. These can be obtained from the [Trello Developer Page](https://trello.com/power-ups/admin/). You can store these in a `.env` file or set them as environment variables.

- **AWS Credentials** (optional for S3 uploads): Set up your AWS credentials through the AWS CLI or environment variables for S3 uploads.

### Example `.env` File:

```
TRELLO_API_KEY=your_api_key
TRELLO_ACCESS_TOKEN=your_access_token
AWS_ACCESS_KEY_ID=your_aws_access_key
AWS_SECRET_ACCESS_KEY=your_aws_secret_key
AWS_REGION=your_aws_region
```

## Usage

### 1. Export Trello Board Data

Run the following command to start the exporter:

```bash
trello-csv
```

### 2. Specify Output Directory

You can specify where to save the generated file:

- **Save Locally** (to a specific directory):

```bash
trello-csv --output-dir ./csv
```

If the `--output-dir` option is omitted, the file will be saved to your system’s default Downloads folder:

- **macOS/Linux**: `~/Downloads`
- **Windows**: `C:\Users\YourUsername\Downloads`

- **Upload to AWS S3**:

```bash
trello-csv --output-dir s3://your-bucket-name/path/to/directory/
```

### 3. Specify AWS Profile (optional)

If you use multiple AWS profiles, you can specify the profile to use:

```bash
trello-csv --aws-profile my-aws-profile
```

## Example Command

1. **Save file locally**:

```bash
trello-csv --output-dir ./trello_exports
```

2. **Upload to AWS S3**:

```bash
trello-csv --output-dir s3://my-bucket/trello_exports --aws-profile my-aws-profile
```

This will export your Trello board data to `./trello_exports` or directly to the S3 bucket `my-bucket/trello_exports`.

## Development (Optional)

If you want to contribute or build the project locally, here are the steps:

### 1. Clone the repository:

```bash
git clone https://github.com/your-username/trello-csv-exporter.git
cd trello-csv-exporter
```

### 2. Set up the virtual environment and install dependencies:

```bash
python -m venv venv
source venv/bin/activate  # On Linux/macOS
.\venv\Scripts\activate   # On Windows
pip install -r requirements.txt
```

### 3. Build the package:

To build the package locally, run:

```bash
make install
```

## License

This project is licensed under the MIT License - see the [LICENSE](./LICENSE) file for details.
