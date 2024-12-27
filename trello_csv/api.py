import os

import boto3
import requests
from botocore.exceptions import NoCredentialsError
from botocore.utils import urlparse
from loguru import logger

from trello_csv.excel import create_excel_sheet


def get_trello_boards(api_key, access_token):
    url = "https://api.trello.com/1/members/me/boards"

    params = {
        "key": api_key,
        "token": access_token,
    }

    response = requests.get(url, params=params)

    if response.status_code == 200:
        boards = response.json()
        return {board["name"]: board["id"] for board in boards}
    else:
        logger.error(
            f"Failed to fetch Trello boards. Status code: {response.status_code}"
        )
        return None


def get_trello_board_data(api_key, access_token, board_id):
    url = f"https://api.trello.com/1/boards/{board_id}"

    params = {
        "key": api_key,
        "token": access_token,
        "cards": "open",
        "lists": "all",
        "labels": "all",
    }

    response = requests.get(url, params=params)

    if response.status_code == 200:
        return response.json()
    else:
        logger.error(f"Failed to fetch board data. Status code: {response.status_code}")
        return None


def upload_to_s3(data_to_export, board_name, s3_url, aws_profile=None):
    s3_url_parts = urlparse(s3_url)
    bucket_name = s3_url_parts.netloc

    sanitized_board_name = "".join(
        c for c in board_name if c.isalnum() or c in (" ", "_")
    )
    file_path = create_excel_sheet(data_to_export, board_name, s3_url)

    if not file_path:
        logger.error("Error creating the Excel file.")
        return False

    try:
        file_key = os.path.join(
            s3_url_parts.path.lstrip("/"),
            f"{sanitized_board_name}_trello_template.xlsx",
        )
        session = boto3.Session(profile_name=aws_profile)
        s3 = session.client("s3")
        s3.upload_file(file_path, bucket_name, file_key)
        logger.info(f"File uploaded successfully to S3: {s3_url}")
        return True

    except NoCredentialsError:
        logger.error(
            "AWS credentials not available or not authorized. Please check your credentials."
        )
        return False
    except Exception as e:
        logger.error(f"Error uploading to S3: {str(e)}")
        return False
    finally:
        if os.path.exists(file_path):
            os.remove(file_path)
            logger.info(f"Temporary file {file_path} deleted")
