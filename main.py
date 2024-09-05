from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import googleapiclient.http

import os
import io
import sys

SCOPES = [
  "https://www.googleapis.com/auth/spreadsheets",
  'https://www.googleapis.com/auth/drive'
]

SPREADSHEET_ID = "__YOUR__SHEET__ID__"

def collect_markets(path_file=None) -> set:
  if path_file == None:
    print('No path_file given')
    return None
  else:
    markets = set()
    with open(path_file, 'r') as markets_file:
      markets_list = markets_file.read().split('\n')
      for market in markets_list:
        markets.add(market)
      markets_file.close()
    print(f'Total Markets collected: {len(markets)}')
    return markets


def main():
  credentials = None
  if os.path.exists('token.json'):
    credentials = Credentials.from_authorized_user_file('token.json', SCOPES)
  if not credentials or not credentials.valid:
    if credentials and credentials.expired and credentials.refresh_token:
      credentials.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
      credentials = flow.run_local_server(port=0)
    with open('token.json', 'w') as token:
      token.write(credentials.to_json())
  try:
    service = build('sheets', 'v4', credentials=credentials)
    sheets = service.spreadsheets()

    # Collecting Markets
    try:
      markets = collect_markets(sys.argv[1])
    except IndexError as ie:
      print(f'Usage: {sys.argv[0]} marktes_file.txt')
      return

    for market in markets:
      try:
        sheets.values().update(
          spreadsheetId=SPREADSHEET_ID,
          range=f"Sheet1!A1",
          valueInputOption='USER_ENTERED',
          body={
            "values": [[f'=GOOGLEFINANCE("{market}", "Close", DATE(2006, 5, 9), DATE(2024, 5, 9))']]
          }
        ).execute()

        # Exporting
        result = sheets.values().get(spreadsheetId=SPREADSHEET_ID, range='Sheet1!A1:B4327').execute()

        values = result.get('values', [])

        print(f'Collecting time series from {market}')
        with open(f'./Markets/{market}.csv', 'w') as csv_file:
          for row in values:
            csv_file.write(f'{row[0]},{row[1]}\n')
          csv_file.close()
      except:
        print(f'Error processing time series from {market}')
        os.system(f'del .\\Datasets\\{market}.csv')

  except HttpError as error:
    print(error)

if __name__ == "__main__":
  main()
