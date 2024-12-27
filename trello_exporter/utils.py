import os
import sys
import time

import boto3
from botocore.exceptions import NoCredentialsError
from colorama import Fore, Style, init
from dotenv import load_dotenv
from loguru import logger

from .api import get_trello_boards

init(autoreset=True)


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
            print(f"{Fore.GREEN}Available Trello boards:{Style.RESET_ALL}")
            for index, (board_name, _) in enumerate(boards.items(), start=1):
                print(f"{Fore.CYAN}{index}. {board_name}{Style.RESET_ALL}")

            selection = input(
                f"{Fore.YELLOW}Enter the number of the board you want to export: {Style.RESET_ALL}"
            )
            print()

            try:
                selection = int(selection)
                if 1 <= selection <= len(boards):
                    selected_board = list(boards.values())[selection - 1]
                    return selected_board
                else:
                    print(
                        f"{Fore.RED}Invalid selection. Please enter a valid number.{Style.RESET_ALL}"
                    )
                    time.sleep(2)
            except ValueError:
                print(
                    f"{Fore.RED}Invalid input. Please enter a number.{Style.RESET_ALL}"
                )
                time.sleep(2)
        else:
            print(
                f"{Fore.RED}No Trello boards found. Please check your credentials.{Style.RESET_ALL}"
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
