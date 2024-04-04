

# importing some important libraries
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import os


curr_dir = os.path.dirname(os.path.realpath(__file__))
key_file = os.path.join(curr_dir, 'my.json')

# authorizing scopes
scopes = ["https://www.googleapis.com/auth/drive"]

# creating credentials
credentials = Credentials.from_service_account_file(key_file, scopes=scopes)

# initializing resource object
resource = build('sheets', 'v4', credentials=credentials)

# spreadsheet_id
spreadsheet_id = "1oTN75OhcvWqNHHtUxjfvQBXI3Y_c-g7QnvDzTXrbdSU"

# workbook
spreadsheets = resource.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
for sheet in spreadsheets['sheets']:
    print(f"Title: {sheet['properties']['title']}, ID: {sheet['properties']['sheetId']}")


# taking sheet id as input
sheet_title = input("Enter Sheet title: ")
sheet_range = f"{sheet_title}!F5:K5"

# writing to a sheet
values = [1, 2, 3, "salman", 5, 3, 1, "ahmad"]
workbook = resource.spreadsheets()
request = workbook.values().append(
    spreadsheetId=spreadsheet_id,
    range=sheet_range,
    valueInputOption="USER_ENTERED",
    insertDataOption="INSERT_ROWS",
    body={
        "values": [values]
    }
)
response = request.execute()
print(response)

