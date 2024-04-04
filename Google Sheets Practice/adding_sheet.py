

# importing some important libraries
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import os
from pprint import pprint


curr_dir = os.path.dirname(os.path.realpath(__file__))
key_file = os.path.join(curr_dir, 'my.json')

# creating scopes
scopes = ['https://www.googleapis.com/auth/drive']

# spreadsheet_id
spreadsheet_id = "1oTN75OhcvWqNHHtUxjfvQBXI3Y_c-g7QnvDzTXrbdSU"

# initializing credentials
credentials = Credentials.from_service_account_file(key_file, scopes=scopes)

# creating service object
service = build('sheets', 'v4', credentials=credentials)

sheet = service.spreadsheets()

# adding a new sheet to the spreadsheet
new_sheet_title = "sheet23"
request = sheet.batchUpdate(spreadsheetId=spreadsheet_id,
                            body={
                                "requests": [{
                                    "addSheet": {
                                        "properties": {"title": new_sheet_title}
                                    }
                                }]
                            })
response = request.execute()
print(response)

