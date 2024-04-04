

# importing some important libraries
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import os


curr_dir = os.path.dirname(os.path.realpath(__file__))
key_file = os.path.join(curr_dir, 'my.json')

# spreadsheet_id
spreadsheet_id = "1oTN75OhcvWqNHHtUxjfvQBXI3Y_c-g7QnvDzTXrbdSU"

# authorizing scopes
scopes = ['https://www.googleapis.com/auth/drive']

# creating credentials
credentials = Credentials.from_service_account_file(key_file, scopes=scopes)

# initializing resource object
resource = build('sheets', 'v4', credentials=credentials)

sheets = resource.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
for sheet in sheets['sheets']:
    print(f"Sheet Title: {sheet['properties']['title']}, Sheet ID: {sheet['properties']['sheetId']}")

sheet_title = input("Enter sheet title: ")
sheet_range = f"{sheet_title}!A4:F78"

# reading data from sheet for the required range
workbook = resource.spreadsheets()
request = workbook.values().get(
    spreadsheetId=spreadsheet_id,
    range=sheet_range,
)
response = request.execute()
values = response.get("values", [])
for value in values:
    print(value)
