# api.py

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

def upload_to_s3(data_to_export, board_name, s3_url, aws_profile=None):
    s3_url_parts = urlparse(s3_url)
    bucket_name = s3_url_parts.netloc
    sanitized_board_name = "".join(c for c in board_name if c.isalnum() or c in (' ', '_'))
    sanitized_filename = f'{sanitized_board_name}_trello_template.xlsx'

    try:
        with open(sanitized_filename, 'wb') as data_file:
            create_excel_sheet(data_to_export, board_name, data_file)

        session = boto3.Session(profile_name=aws_profile)
        s3 = session.client('s3')
        s3.upload_file(sanitized_filename, bucket_name, s3_url_parts.path.lstrip('/'))
        print("Upload successful")
    except NoCredentialsError:
        print("AWS credentials not available or not authorized. Please check your credentials.")
    except Exception as e:
        print(f"Error uploading to S3: {str(e)}")
    finally:
        os.remove(sanitized_filename)
