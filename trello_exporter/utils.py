import os
import time

from dotenv import load_dotenv
from loguru import logger

from .api import get_trello_boards


def load_environment_variables():
    load_dotenv()
    
    api_key = os.getenv("TRELLO_API_KEY")
    access_token = os.getenv("TRELLO_TOKEN")

    if not api_key or not access_token:
        logger.error("Trello API key or access token is missing. Please check your .env file.")
        sys.exit(1)

    return {'api_key': api_key, 'access_token': access_token}

def select_trello_board(api_key, access_token):
    while True:
        boards = get_trello_boards(api_key, access_token)

        if boards:
            print("Available Trello boards:")
            for index, (board_name, board_id) in enumerate(boards.items(), start=1):
                print(f"{index}. {board_name}")

            selection = input("Enter the number of the board you want to export: ")
            print()

            try:
                selection = int(selection)
                if 1 <= selection <= len(boards):
                    selected_board = list(boards.values())[selection - 1]
                    return selected_board
                else:
                    print("Invalid selection. Please enter a valid number.")
                    time.sleep(2)
            except ValueError:
                print("Invalid input. Please enter a number.")
                time.sleep(2)
        else:
            print("No Trello boards found. Please check your credentials")
            sys.exit(1)

def extract_card_data(board_data):
    data_to_export = []
    
    board_name = board_data['name'] 

    for card in board_data['cards']:
        list_name = None
        for list in board_data['lists']:
            if list['id'] == card['idList']:
                list_name = list['name']
                break

        if list_name != "Maintenance":
            card_data = {
                'Name': card['name'],
                'Description': card['desc'],
                'List': list_name,
                'Labels': ', '.join([label['name'] for label in card['labels']])
            }
            data_to_export.append(card_data)

    return data_to_export, board_name
