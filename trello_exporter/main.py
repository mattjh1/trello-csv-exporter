import os

from trello_exporter.api import get_trello_board_data, get_trello_boards
from trello_exporter.excel import create_excel_sheet
from trello_exporter.utils import (extract_card_data,
                                   load_environment_variables,
                                   select_trello_board)


def main():
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

            create_excel_sheet(card_data, board_name)
            logger.info("Excel file populated successfully")

            logger.info("Script completed")


if __name__ == "__main__":
    main()
