import requests
from bs4 import BeautifulSoup as bs
import lxml
import pickle
from loguru import logger
import datetime
import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
import pickle
import time
import pars_Dohod

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



def clear_area():
    """
    Clear cell in area
    """
    service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range='AP3:AS21',
    ).execute()


def read_ticker():
    """
    Читаем колонку F и с этими данным топаем на Доход
    """
    # Read ticket for RBK, column -D-
    values = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range='F3:F21',
        majorDimension='COLUMNS'
    ).execute()
    lst = []
    for i in values['values'][0]:
        lst.append(str(i).lower())

    with open(".\data_pkl\Tiker_F_collum.pkl", 'wb') as f:
        pickle.dump(lst, f)


def get_data(tiker: str):
    url = f'https://www.dohod.ru/ik/analytics/dividend/{tiker}'
    headers = {'accept': '*/*',
               'User-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:67.0) Gecko/20100101  Firefox/67.0'}

    r = requests.get(url=url, headers=headers)
    lst = []
    if r.status_code == 200:
        soup = bs(r.text, "lxml")
        par = soup.find('p').stripped_strings

        for i in par:
            lst.append(i)

        lst = [lst[1], lst[7], lst[10][:-1].replace('.',','), lst[12]]

        return lst

    else:
        print(f"er {r.status_code}")
        lst = ['-', '-', '-', '-']
        return lst


def send_dohod(index_cell: int, lst: list):
    # Пример записи в файл

    values = service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": [
                {"range": f"AP{index_cell}:AS{index_cell}",
                 "majorDimension": "ROWS",
                 "values": [
                     [
                         f'{lst[0]}', f'{lst[1]}', f'{lst[2]}', f'{lst[3]}'  # Актуальная инфа

                     ],
                 ]
                 },
            ]
        }
    ).execute()
    logger.info(f"{index_cell} row sended")


def main():
    clear_area()
    read_ticker()
    with open(".\data_pkl\Tiker_F_collum.pkl", 'rb') as f:
        lst = pickle.load(f)
    cell = 3
    for i in lst:
        data = get_data(i)
        send_dohod(cell, data)
        cell += 1
        time.sleep(1)


if __name__ == '__main__':
    main()
