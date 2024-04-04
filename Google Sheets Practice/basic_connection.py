

# importing some important libraries
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import os


curr_dir = os.path.dirname(os.path.realpath(__file__))
key_file = os.path.join(curr_dir, "my.json")

# creating scopes
scopes = ["https://www.googleapis.com/auth/drive"]

# spreadsheet_id
spreadsheet_id = "1oTN75OhcvWqNHHtUxjfvQBXI3Y_c-g7QnvDzTXrbdSU"

# setting up credentials
credentials = Credentials.from_service_account_file(key_file, scopes=scopes)

# initializing service object
service = build('sheets', 'v4', credentials=credentials)

# now we can create a basic object for future use
spreadsheet = service.spreadsheets()

# printing sheet properties
sheet_properties = spreadsheet.get(spreadsheetId=spreadsheet_id).execute()
print(sheet_properties)
