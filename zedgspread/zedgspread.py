
# This file contains the ZED G Spread class

# importing some important libraries
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import os


class ZEDGSpread:
    """
    A class to automate Google Spread Sheets and to perform various tasks on it.
    """
    
    __credentials = None
    __key_file = None
    __scopes = None
    __resource = None
    __spreadsheet_id = None
    __workbook = None

    # utility functions
    def __check_sheet_exists(self, sheet_name: str):
        """
        Checks if the sheet is present in the workbook or not.
        :param sheet_name:
        :return:
        """

        # getting all sheets from workbook
        sheets = self.list_sheet_names()[0]
        if sheet_name in sheets:
            return True
        else:
            return False

    # Constructor
    def __init__(self, path_to_secrets_file: str, spreadsheet_id: str, scopes: [str]):
        """
        Initializes KEYFILE and creates a Credential for class to be used.
        :param path_to_secrets_file: 
        :param spreadsheet_id: 
        :param scopes: 
        """

        # initializing key_file
        self.__key_file = path_to_secrets_file
        # initializing spreadsheets_id
        self.__spreadsheet_id = spreadsheet_id
        # authorizing scopes
        self.__scopes = scopes

        # calling some functions to make it work
        self.authorize_credentials()
        self.initialize_resource()
        self.__workbook = self.__resource.spreadsheets()

    # Authorize Credentials
    def authorize_credentials(self):
        """
        Creates Credentials for the Google Sheets and returns them.
        :return:
        """

        # creating credentials
        self.__credentials = Credentials.from_service_account_file(self.__key_file, scopes=self.__scopes)

        # returning
        return self.__credentials

    # Initialize Resource
    def initialize_resource(self):
        """
        Initializes the resource object.
        :return:
        """
        self.__resource = build("sheets", "v4", credentials=self.__credentials)

    # Get Sheets in workbook
    def get_sheets_in_workbook(self):
        """
        Returns the sheets present inside a workbook or a spreadsheet.
        :return:
        """

        sheets = self.__resource.spreadsheets().get(spreadsheetId=self.__spreadsheet_id).execute()
        return sheets

    # List Sheet Name
    def list_sheet_names(self):
        """
        Returns Sheet Title and Sheet ID's currently present in the workbook.
        :return:
        """

        sheets = self.__workbook.get(spreadsheetId=self.__spreadsheet_id).execute()
        sheet_title = []
        sheet_id = []
        for sheet in sheets['sheets']:
            sheet_title.append(sheet['properties']['title'])
            sheet_id.append(sheet['properties']['sheetId'])

        return sheet_title, sheet_id

    # Create Sheet
    def create_sheet(self, sheet_name: str, sheet_index = 1):
        """
        Creates a new sheet and adds it to the workbook.
        :param sheet_index:
        :param sheet_name:
        :return:
        """
        # first, we need to check that sheet doesn't exist already
        sheets = self.list_sheet_names()[0]
        if sheet_name not in sheets:
            self.__workbook.batchUpdate(
                spreadsheetId=self.__spreadsheet_id,
                body={
                    "requests": [{
                        "addSheet": {
                            "properties": {
                                "title": sheet_name,
                                "index": sheet_index
                            }
                        }
                    }]
                }
            ).execute()
        else:
            raise NameError("Sheet already exists.")

    # Delete Sheet
    def delete_sheet(self, sheet_id: int):
        """
        Deletes a sheet from the workbook.
        :param sheet_id:
        :return:
        """

        # first, we need to check that sheet is present in the workbook
        sheet_ids = [int(i) for i in self.list_sheet_names()[1]]
        if sheet_id not in sheet_ids:
            raise NameError("Sheet ID not present.")
        else:
            self.__workbook.batchUpdate(
                spreadsheetId=self.__spreadsheet_id,
                body={
                    "requests": [{
                        "deleteSheet": {
                            "sheetId": sheet_id
                        }
                    }]
                }
            ).execute()

    # Read Sheet Content
    def read_sheet_content(self, sheet_name=None, sheet_range=None):
        """
        Reads the specified range from the given sheet title and
        returns the contents.
        :param sheet_name:
        :param self:
        :param sheet_range:
        :return:
        """

        # a check for sheet_name is not none
        if sheet_name is None:
            raise KeyError("Sheet name cannot be none.")

        # [CHECK] if sheet is present in workbook or not
        if not self.__check_sheet_exists(sheet_name):
            raise KeyError("Sheet does not exist.")

        if sheet_range is None:
            workbook_range = sheet_name
        else:
            workbook_range = f"{sheet_name}!{sheet_range}"

        request = self.__workbook.values().get(
            spreadsheetId=self.__spreadsheet_id,
            range=workbook_range
        )
        response = request.execute()
        contents = response.get("values", [])
        return contents

    # Write To Sheet
    def write_to_sheet(self, sheet_name=None, sheet_range=None, contents=None):
        """
        Writes the contents to the given sheet at specified range.
        :param contents:
        :param sheet_name:
        :param sheet_range:
        :return:
        """

        # check if sheet name is none
        if sheet_name is None:
            raise KeyError("Sheet name cannot be none.")

        # checking is sheet_name is already present in workbook or not
        if not self.__check_sheet_exists(sheet_name):
            raise KeyError("Sheet does not exist.")

        if sheet_range is None:
            workbook_range = sheet_name
        else:
            workbook_range = f"{sheet_name}!{sheet_range}"

        # check for values
        if all(isinstance(v, list) for v in contents):
            values = contents
        else:
            values = [contents]

        # writing to sheet
        request = self.__workbook.values().append(
            spreadsheetId=self.__spreadsheet_id,
            range=workbook_range,
            valueInputOption="USER_ENTERED",
            insertDataOption="OVERWRITE",
            body={
                "values": values
            }
        )
        response = request.execute()
        return response

    # Clear Sheet Content
    def clear_sheet(self, sheet_name=None, sheet_range=None):
        """
        Clears the whole contents of the given sheet at specified range.
        If range is None, then it clears the whole sheet.
        :param sheet_range:
        :param sheet_name:
        :return:
        """

        # checking if sheet_name is None
        if sheet_name is None:
            raise KeyError("Sheet name cannot be none.")

        # [CHECK] if sheet name is present in workbook or not
        if not self.__check_sheet_exists(sheet_name):
            raise KeyError("Sheet does not exist.")

        if sheet_range is None:
            workbook_range = sheet_name
        else:
            workbook_range = f"{sheet_name}!{sheet_range}"

        # Now clearing as specified in the range
        request = self.__workbook.values().clear(
            spreadsheetId=self.__spreadsheet_id,
            range=workbook_range,
            body={}
        )
        response = request.execute()
        return response

    # Copy Sheet
    def copy_sheet(self, sheet1=None, sheet2=None, sheet_range=None):
        """
        Copy the contents of sheet2 to the sheet1.
        :param sheet2:
        :param sheet1:
        :param sheet_range:
        :return:
        """

        # [CHECK] if both sheet names are not none
        if sheet1 is None or sheet2 is None:
            raise NameError("Sheet name cannot be none.")

        # [CHECK] if both sheets exists or not
        if not self.__check_sheet_exists(sheet1):
            raise KeyError(f"f{sheet1} does not exist.")

        if not self.__check_sheet_exists(sheet2):
            raise KeyError(f"{sheet2} does not exist.")

        # we need to copy contents of sheet 2 to the sheet1
        values = self.read_sheet_content(sheet2)

        # now, writing them to the sheet 1
        # before writing we need to clear sheet
        self.clear_sheet(sheet1)
        self.write_to_sheet(sheet_name=sheet1, contents=values)

    # Dataframe to Sheet
    def dataframe_to_sheet(self, dataframe, sheet_name=None):
        """
        Converts the given dataframe to the spreadsheet.
        :param sheet_name:
        :param dataframe:
        :return:
        """

        # [CHECK] sheet name cannot be none
        if sheet_name is None:
            raise NameError("Sheet Name cannot be none.")

        # [CHECK] creates a new sheet if it is not present in the workbook
        if not self.__check_sheet_exists(sheet_name):
            self.create_sheet(sheet_name)
        else:
            # clear contents of the sheet from the workbook
            self.clear_sheet(sheet_name=sheet_name)

        # writing columns
        self.write_to_sheet(sheet_name=sheet_name, contents=dataframe.columns.tolist())
        # writing values
        jinga = dataframe.values.tolist()
        self.write_to_sheet(sheet_name=sheet_name, contents=jinga)
