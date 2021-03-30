import datetime
import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials

# Файл, полученный в Google Developer Console
CREDENTIALS_FILE = 'creds.json'
# ID Google Sheets документа (можно взять из его URL)
spreadsheet_id = '1M74wp0CI6CgOpY7MB1grFAxRJPlKpIS2f3H6Kcu1fHI'

# Авторизуемся и получаем service — экземпляр доступа к API
credentials = ServiceAccountCredentials.from_json_keyfile_name(
    CREDENTIALS_FILE,
    ['https://www.googleapis.com/auth/spreadsheets',
     'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)


def send_time(cell):
    cell += 3
    today = datetime.datetime.today()
    print(today)

    update_data = today.strftime("%d.%m.%Y %H.%M")  # 2017-04-05-00.18.00

    values = service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": [
                {"range": f"A{cell}", #:B{cell}",
                 "majorDimension": "ROWS",
                 "values": [
                     [
                         f'{update_data}'  # Актуальная инфа

                     ],
                 ]
                 },
            ]
        }
    ).execute()


if __name__ == '__main__':
    send_time()
