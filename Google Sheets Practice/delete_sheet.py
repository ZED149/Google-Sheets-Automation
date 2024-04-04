
# importing some important libraries
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import os
from pprint import pprint

curr_dir = os.path.dirname(os.path.realpath(__file__))
key_file = os.path.join(curr_dir, 'my.json')

# setting up scopes
scopes = ["https://www.googleapis.com/auth/drive"]

# spreadsheet_id
spreadsheet_id = "1oTN75OhcvWqNHHtUxjfvQBXI3Y_c-g7QnvDzTXrbdSU"

# creating credentials
credentials = Credentials.from_service_account_file(key_file, scopes=scopes)

# initializing resource object
service = build('sheets', 'v4', credentials=credentials)

# printing spreadsheet properties
spreadsheet_properties = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
for sheet in spreadsheet_properties['sheets']:
    print(sheet)

# sheet
sheet = service.spreadsheets()

# request object
sheetId = int(input("Enter sheet id to delete sheet: --> "))
request = sheet.batchUpdate(spreadsheetId=spreadsheet_id,
                              body={
                                  "requests": [{
                                      "deleteSheet": {
                                          "sheetId": sheetId
                                      }
                                  }]
                              })
response = request.execute()
print(response)
