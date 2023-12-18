import requests

def get_trello_boards(api_key, access_token):
    url = "https://api.trello.com/1/members/me/boards"

    params = {
        'key': api_key,
        'token': access_token,
    }

    response = requests.get(url, params=params)

    if response.status_code == 200:
        boards = response.json()
        return {board['name']: board['id'] for board in boards}
    else:
        print(f'Failed to fetch Trello boards. Status code: {response.status_code}')
        return None

def get_trello_board_data(api_key, access_token, board_id):
    url = f"https://api.trello.com/1/boards/{board_id}"

    params = {
        'key': api_key,
        'token': access_token,
        'cards': 'all',
        'lists': 'all',
        'labels': 'all',
    }

    response = requests.get(url, params=params)

    if response.status_code == 200:
        return response.json()
    else:
        print(f'Failed to fetch board data. Status code: {response.status_code}')
        return None
