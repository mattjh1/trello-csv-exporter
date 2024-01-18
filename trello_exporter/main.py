import argparse
import sys

import boto3
from loguru import logger

from trello_exporter.api import get_trello_board_data, upload_to_s3
from trello_exporter.excel import create_excel_sheet
from trello_exporter.utils import (
    check_aws_credentials,
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
    aws_profile = args.aws_profile

    if is_s3 and not aws_profile:
        aws_profile = boto3.DEFAULT_SESSION.profile_name

    if is_s3 and not check_aws_credentials(aws_profile):
        logger.error(
            "AWS credentials not found or not authorized. Please check your credential"
        )
        sys.exit(1)

    credentials = load_environment_variables()
    api_key = credentials["api_key"]
    access_token = credentials["access_token"]
    output_dir = args.output_dir or "./csv"


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
