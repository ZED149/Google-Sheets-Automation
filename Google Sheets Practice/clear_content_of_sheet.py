

# importing some important libraries
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import os


curr_dir = os.path.dirname(os.path.realpath(__file__))
key_file = os.path.join(curr_dir, "my.json")

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

# taking sheet title as input
sheet_title = input("Enter sheet title: ")
sheet_range = f"{sheet_title}"

# clearing content of the sheet
workbook = resource.spreadsheets()
request = workbook.values().clear(
    spreadsheetId=spreadsheet_id,
    range=sheet_range,
    body={}
)
response = request.execute()
print(response)
