import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name('bikroy-4cad724c2b4d.json', scope)

gc = gspread.authorize(credentials)

wks = gc.open('bikroy.com').sheet1

me= wks.get_all_records()

print(me)