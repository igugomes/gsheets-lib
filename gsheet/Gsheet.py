import os
import logging

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse

    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    import argparse

    flags = None

__author__ = "Igor Gomes"
__email__ = "igor.mlgomes@gmail.com"

"""
BEFORE RUNNING:
---------------
1. If not already done, enable the Google Sheets API
   and check the quota for your project at
   https://console.developers.google.com/apis/api/sheets
2. Install the Python client library for Google APIs by running
   `pip install --upgrade google-api-python-client`
"""

SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets Library'


class Gsheet(object):
    """
    Implementation of methods to acces the Google Sheets API
    """

    def __init__(self, spreadsheet_id):
        """
        Constructor of the class
        :param spreadsheet_id: The ID of the spreadsheet to update.
        """
        self._spreadsheet_id = spreadsheet_id

    @staticmethod
    def get_credentials(credential):
        """Gets valid user credentials from storage.

        If nothing has been stored, or if the stored credentials are invalid,
        the OAuth2 flow is completed to obtain the new credentials.

        Returns:
            Credentials, the obtained credential.
        """
        workspace = os.path.abspath('.')
        credential_dir = os.path.join(workspace, '.credentials')
        if not os.path.exists(credential_dir):
            os.makedirs(credential_dir)
        credential_path = os.path.join(credential_dir,
                                       credential)
        credentials = ''
        store = Storage(credential_path)
        try:
            credentials = store.get()
        except RuntimeError as err:
            logging.info('Working with flow-based credentials instantiation')
            logging.error('Error message: \n{}'.format(err))

        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets(credential_path, SCOPES)
            flow.user_agent = APPLICATION_NAME
            if flags:
                credentials = tools.run_flow(flow, store, flags)
            else:  # Needed only for compatibility with Python 2.6
                credentials = tools.run(flow, store)
            logging.info('Storing credentials to ' + credential_path)
        return credentials

    def get_service(self):
        return discovery.build('sheets', 'v4', credentials=self.get_credentials(CLIENT_SECRET_FILE))

    def sheet_append_values(self, range_, values):
        """
        Appends the given values to the given spreadsheet and range.
        :param range_: The A1 notation of a range to search for a logical table of data.
        :param values: The values to append to the spreadsheet
        :return: The append action
        """
        # How the input data should be interpreted.
        value_input_option = 'RAW'

        # How the input data should be inserted.
        insert_data_option = 'INSERT_ROWS'
        value_range_body = {'values': values}

        request = self.get_service().spreadsheets().values().append(spreadsheetId=self._spreadsheet_id, range=range_,
                                                                    valueInputOption=value_input_option,
                                                                    insertDataOption=insert_data_option,
                                                                    body=value_range_body)
        return request.execute()

    def sheet_clear_values(self, range_):
        """
        Clear the given values for the given spreadsheet and range
        :param range_: The A1 notation of the values to clear.
        :return: The clear action
        """
        request = self.get_service().spreadsheets().values().clear(spreadsheetId=self._spreadsheet_id, range=range_)
        return request.execute()

    def sheet_get_values(self, range_, value_render_option=""):
        """
        Gets the values for the given spreadsheet and range
        :param range_: The ID of the spreadsheet to retrieve data from.
        :param value_render_option: How values should be represented in the output.
        :return: The values catch related to the given arguments
        """
        # The default render option is ValueRenderOption.FORMATTED_VALUE.
        if 'un' in value_render_option:
            value_render_option = 'UNFORMATTED_VALUE'
        elif 'formula' in value_render_option:
            value_render_option = 'FORMULA'
        else:
            value_render_option = 'FORMATTED_VALUE'

        # How dates, times, and durations should be represented in the output.
        # This is ignored if value_render_option is
        # FORMATTED_VALUE.
        # The default dateTime render option is [DateTimeRenderOption.SERIAL_NUMBER].
        date_time_render_option = 'SERIAL_NUMBER'

        request = self.get_service().spreadsheets().values().get(spreadsheetId=self._spreadsheet_id, range=range_,
                                                                 valueRenderOption=value_render_option,
                                                                 dateTimeRenderOption=date_time_render_option)
        return request.execute()

    def sheet_update_values(self, range_, values):
        """
        Updates the given values for the given spreadsheet and range
        :param range_: The A1 notation of the values to update.
        :param values: The values to be updated in the spreadsheet
        :return: The updated values related to teh given arguments
        """
        # How the input data should be interpreted.
        value_input_option = 'RAW'

        value_range_body = {
            'values': values
        }

        request = self.get_service().spreadsheets().values().update(spreadsheetId=self._spreadsheet_id, range=range_,
                                                                    valueInputOption=value_input_option,
                                                                    body=value_range_body)
        return request.execute()
