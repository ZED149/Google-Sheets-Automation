

# importing some important libraies
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

title = "Python Google Sheets Automation"

# workbook
workbook = resource.spreadsheets()
request = workbook.batchUpdate(
    spreadsheetId=spreadsheet_id,
    body={
        "requests": [{
            "updateSpreadsheetProperties": {
                "properties": {
                    "title": title
                }, "fields": "title"
            }
        }]
    }
)
response = request.execute()
print(response)
