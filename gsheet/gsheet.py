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

import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets Library'


def get_credentials(credential):
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.abspath('.')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   credential)

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


credentials = get_credentials('')

service = discovery.build('sheets', 'v4', credentials=credentials)


def sheet_append_values(spreadsheet_id, range_, values):
    """
    Appends the given values to the given spreadsheet and range.
    :param spreadsheet_id: The ID of the spreadsheet to update.
    :param range_: The A1 notation of a range to search for a logical table of data.
    :param values: The values to append to the spreadsheet
    :return: The append action
    """
    spreadsheet_id = spreadsheet_id

    # Values will be appended after the last row of the table.
    range_ = range_

    # How the input data should be interpreted.
    value_input_option = 'RAW'

    # How the input data should be inserted.
    insert_data_option = 'OVERWRITE'
    value_range_body = {
        'values': values
    }
    request = service.spreadsheets().values().append(spreadsheetId=spreadsheet_id, range=range_,
                                                     valueInputOption=value_input_option,
                                                     insertDataOption=insert_data_option,
                                                     body=value_range_body)
    return request.execute()


def sheet_clear_values(spreadsheet_id, range_, values):
    """
    Clear the given values for the given spreadsheet and range
    :param spreadsheet_id: The ID of the spreadsheet to update.
    :param range_: The A1 notation of the values to clear.
    :param values: The values to be cleared from the spreadsheet
    :return: The clear action
    """
    spreadsheet_id = spreadsheet_id
    range_ = range_

    clear_values_request_body = {
        'values': values
    }

    request = service.spreadsheets().values().clear(spreadsheetId=spreadsheet_id, range=range_,
                                                    body=clear_values_request_body)
    return request.execute()


def sheet_get_values(spreadsheet_id, range_, value_render_option=None):
    """
    Gets the values for the given spreadsheet and range
    :param spreadsheet_id: The ID of the spreadsheet to retrieve data from.
    :param range_: The ID of the spreadsheet to retrieve data from.
    :param value_render_option: How values should be represented in the output.
    :return: The values catch related to the given arguments
    """
    spreadsheet_id = spreadsheet_id
    range_ = range_

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

    request = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_,
                                                  valueRenderOption=value_render_option,
                                                  dateTimeRenderOption=date_time_render_option)
    return request.execute()


def sheet_update_values(spreadsheet_id, range_, values):
    """
    Updates the given values for the given spreadsheet and range
    :param spreadsheet_id: The ID of the spreadsheet to update.
    :param range_: The A1 notation of the values to update.
    :param values: The values to be updated in the spreadsheet
    :return: The updated values related to teh given arguments
    """
    spreadsheet_id = spreadsheet_id
    range_ = range_

    # How the input data should be interpreted.
    value_input_option = 'RAW'

    value_range_body = {
        'values': values
    }

    request = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=range_,
                                                     valueInputOption=value_input_option, body=value_range_body)
    return request.execute()
