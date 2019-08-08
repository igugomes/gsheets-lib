"""
BEFORE RUNNING:
---------------
1. If not already done, enable the Google Sheets API
   and check the quota for your project at
   https://console.developers.google.com/apis/api/sheets
2. Install the Python client library for Google APIs by running
   `pip install --upgrade google-api-python-client`
"""
from __future__ import print_function

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


def sheet_append(spreadsheet_id, range, values):
    # The ID of the spreadsheet to update.
    spreadsheet_id = spreadsheet_id
    # The A1 notation of a range to search for a logical table of data.
    # Values will be appended after the last row of the table.
    range_ = range
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
    response = request.execute()
    return response
