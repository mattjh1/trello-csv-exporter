# main.py

import argparse
import os

from trello_exporter.api import get_trello_board_data, get_trello_boards
from trello_exporter.excel import create_excel_sheet
from trello_exporter.utils import (extract_card_data,
                                   load_environment_variables,
                                   select_trello_board)


def main():
    parser = argparse.ArgumentParser(description="Trello Exporter")
    parser.add_argument("--output-dir", help="Output directory path (local or S3)")
    parser.add_argument("--aws-profile", help="AWS profile for S3 (optional)")

    args = parser.parse_args()

    is_s3 = args.output_dir and args.output_dir.startswith("s3://")
    aws_profile = args.aws_profile

    if is_s3 and not aws_profile:
        # Use the default AWS profile if not specified
        aws_profile = boto3.DEFAULT_SESSION.profile_name

    if is_s3 and not check_aws_credentials(aws_profile):
        print("AWS credentials not found or not authorized. Please check your credentials.")
        return

    credentials = load_environment_variables()
    api_key = credentials["api_key"]
    access_token = credentials["access_token"]

    selected_board_id = select_trello_board(api_key, access_token)

    if selected_board_id:
        board_data = get_trello_board_data(api_key, access_token, selected_board_id)

        if board_data:
            logger.info("Starting board CSV exporter")

            card_data, board_name = extract_card_data(board_data)
            logger.info(f"Extracted {len(card_data)} cards")

            output_dir = args.output_dir or "./csv"
            
            if is_s3:
                upload_to_s3(card_data, board_name, output_dir, aws_profile)
            else:
                create_excel_sheet(card_data, board_name, output_dir)


if __name__ == "__main__":
    main()
