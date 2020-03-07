''' 口罩實名制健保藥局查詢

    References:
        Google Sheets API Quickstart:
        https://developers.google.com/sheets/api/quickstart/python

        Google Client Library in Python
        https://developers.google.com/sheets/api/guides/libraries
'''
from __future__ import print_function
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

''' Application dependencies
'''
import os
import pandas as pd
from types import SimpleNamespace

''' Read global configurations
'''
cfg = SimpleNamespace()
try:
    # Load configurations from private module
    import secrets as s
    cfg.SPREADSHEET_ID = s.SPREADSHEET_ID
    cfg.SPREADSHEET_RANGE = s.SPREADSHEET_MASK_UPDATE_RANGE
    cfg.ADDRESS_FILTER = s.DRUGSTORE_ADDRESS_FILTER
except ImportError:
    # Load secrets from environment variables
    cfg.SPREADSHEET_ID = os.environ.get('SPREADSHEET_ID')
    cfg.SPREADSHEET_RANGE = os.environ.get('SPREADSHEET_RANGE')
    cfg.ADDRESS_FILTER = os.environ.get('ADDRESS_FILTER')

''' Google API authentication
'''
def init_google_sheet_api():
    creds = None
    # The file 'token.pickle' stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                        'credentials.json',
                        'https://www.googleapis.com/auth/spreadsheets')
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return build('sheets', 'v4', credentials=creds)

''' Get a data frame with filtered mask supply volume
'''
def get_mask_availability(filter=None):
    df = None
    try:
        url = 'http://data.nhi.gov.tw/Datasets/Download.ashx?rid=A21030000I-D50001-001&l=https://data.nhi.gov.tw/resource/mask/maskdata.csv'
        df = pd.read_csv(url, encoding='utf-8')
        if filter is not None:
            df = df[df['醫事機構地址'].str.match(filter)]
    except ConnectionResetError:
        pass
    return df

if __name__ == '__main__':
    df = get_mask_availability(cfg.ADDRESS_FILTER)
    if df is not None:
        values = df['成人口罩剩餘數'].tolist()
        # Add timestamp, assuming timestamps are the same for all filtered values
        values.insert(0, df['來源資料時間'].tolist()[0])

        service = init_google_sheet_api()
        sheet = service.spreadsheets()
        report = {
            'values': [
                values
            ]
        }
        try:
            result = sheet.values().append(spreadsheetId=cfg.SPREADSHEET_ID,
                                           range=cfg.SPREADSHEET_RANGE,
                                           valueInputOption='USER_ENTERED',
                                           body=report).execute()
        except:
            pass
