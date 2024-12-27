import argparse
import os
import platform
import sys

import boto3
from loguru import logger

from trello_csv.api import get_trello_board_data, upload_to_s3
from trello_csv.excel import create_excel_sheet
from trello_csv.utils import (
    extract_card_data,
    load_environment_variables,
    select_trello_board,
)


def setup():
    global is_s3
    global aws_profile
    global api_key
    global access_token
    global output_dir

    parser = argparse.ArgumentParser(description="Trello CSV Exporter")
    parser.add_argument("--output-dir", help="Output directory path (local or S3)")
    parser.add_argument("--aws-profile", help="AWS profile for S3 (optional)")

    args = parser.parse_args()

    is_s3 = args.output_dir and args.output_dir.startswith("s3://")
    aws_profile = args.aws_profile or None

    if is_s3:
        try:
            if aws_profile:
                session = boto3.Session(profile_name=aws_profile)
            else:
                session = boto3.Session()

            credentials = session.get_credentials()
            if (
                not credentials
                or not credentials.access_key
                or not credentials.secret_key
            ):
                raise ValueError("AWS credentials are incomplete or not loaded.")

            region = session.region_name
            if not region:
                raise ValueError(
                    "AWS region not set. Ensure AWS_REGION or AWS_DEFAULT_REGION is configured."
                )

        except Exception as e:
            logger.error(f"Error initializing AWS session: {str(e)}")
            sys.exit(1)

    credentials = load_environment_variables()
    api_key = credentials["api_key"]
    access_token = credentials["access_token"]

    if args.output_dir:
        output_dir = args.output_dir
    else:
        if platform.system() == "Windows":
            # For Windows, use the user's Downloads folder in their profile
            output_dir = os.path.join(os.environ["USERPROFILE"], "Downloads")
        else:
            # For macOS/Linux, use ~/Downloads
            output_dir = os.path.expanduser("~/Downloads")


def main():
    setup()

    selected_board_id = select_trello_board(api_key, access_token)

    if selected_board_id:
        board_data = get_trello_board_data(api_key, access_token, selected_board_id)

        if board_data:
            logger.info("Starting board CSV exporter")

            card_data, board_name = extract_card_data(board_data)
            logger.info(f"Extracted {len(card_data)} cards")

            success = False
            if is_s3:
                success = upload_to_s3(card_data, board_name, output_dir, aws_profile)
            else:
                success = create_excel_sheet(card_data, board_name, output_dir)

            if success:
                logger.info("Excel file populated successfully")
                logger.info("Script completed")
            else:
                logger.error("Something went wrong during script execution.")


if __name__ == "__main__":
    main()
