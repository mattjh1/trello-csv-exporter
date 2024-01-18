# utils.py

import os
import sys
import time

import boto3
from botocore.exceptions import NoCredentialsError
from dotenv import load_dotenv
from loguru import logger

from .api import get_trello_boards


def check_aws_credentials(profile_name=None):
    try:
        session = boto3.Session(profile_name=profile_name)
        session.client("s3").list_buckets()
        return True
    except NoCredentialsError:
        return False


def load_environment_variables():
    logger.warning(
        "\n\nThe script expects the following env variables to be present and valid:\n---------------------------------\n\tTRELLO_API_KEY\n\tTRELLO_TOKEN\n---------------------------------\n"
    )
    load_dotenv()

    api_key = os.getenv("TRELLO_API_KEY")
    access_token = os.getenv("TRELLO_TOKEN")

    if not api_key or not access_token:
        logger.error(
            "Trello API key or access token is missing. Please check your .env file."
        )
        sys.exit(1)

    return {"api_key": api_key, "access_token": access_token}


def select_trello_board(api_key, access_token):
    while True:
        boards = get_trello_boards(api_key, access_token)

        if boards:
            logger.info("Available Trello boards:", format="{message}")
            for index, (board_name, board_id) in enumerate(boards.items(), start=1):
                logger.info(f"{index}. {board_name}", format="{message}")

            selection = input("Enter the number of the board you want to export: ")
            print()

            try:
                selection = int(selection)
                if 1 <= selection <= len(boards):
                    selected_board = list(boards.values())[selection - 1]
                    return selected_board
                else:
                    logger.error(
                        "Invalid selection. Please enter a valid number.",
                        format="{message)",
                    )
                    time.sleep(2)
            except ValueError:
                logger.error(
                    "Invalid input. Please enter a number.", format="{message)"
                )
                time.sleep(2)
        else:
            logger.error(
                "No Trello boards found. Please check your credentials",
                format="{message)",
            )
            sys.exit(1)


def extract_card_data(board_data):
    data_to_export = []

    board_name = board_data["name"]

    for card in board_data["cards"]:
        list_name = None
        for list in board_data["lists"]:
            if list["id"] == card["idList"]:
                list_name = list["name"]
                break

        if list_name != "Maintenance":
            card_data = {
                "Name": card["name"],
                "Description": card["desc"],
                "List": list_name,
                "Labels": ", ".join([label["name"] for label in card["labels"]]),
            }
            data_to_export.append(card_data)

    return data_to_export, board_name
