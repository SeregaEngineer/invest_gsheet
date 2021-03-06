import requests
from bs4 import BeautifulSoup as bs
import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
import re
import pickle
from loguru import logger
import time

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
        range='AL3:AO26',
    ).execute()
    logger.info('Area cleared')


def get_html_code():

    headers = {'accept': '*/*',
               'User-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:67.0) Gecko/20100101 Firefox/67.0'}
    url = "https://www.dohod.ru/ik/analytics/share"

    r = requests.get(url=url, headers=headers)
    if r.status_code == 200:
        soup = bs(r.text, 'lxml')
        with open(".\data_pkl\data_from_dohod.pkl", 'wb') as f:
            pickle.dump(str(soup), f)
    else:
        logger.info(f"Error server Dohod.ru: {r.status_code}")

    logger.info('The soup is ready')


def read_ticker():
    """
    read ticket from sheet
    """
    # Read name for Dohod, column -C-
    values = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range='C3:C25',
        majorDimension='COLUMNS'
    ).execute()
    print(values)
    with open(".\data_pkl\Ticker_C_collum.pkl", 'wb') as f:
        pickle.dump(values, f)
    logger.info("")

def read_from_file(txt: str):
    lst = []
    with open(".\data_pkl\data_from_dohod.pkl", 'rb') as f:
        soup = bs(pickle.load(f), 'lxml')
    tr = soup.find_all('tr')
    for td in tr:
        # Если нашли эмитента тозабираем данные
        if td.find("td", text=f"{txt}") is not None:
            td_dividend = td.find_all("td", attrs={re.compile("dividend-column")})
            yield_ = td_dividend[0].get_text(strip=True)
            dsi = td_dividend[1].get_text(strip=True)
            strategy = td_dividend[2].get_text(strip=True)
            rating = td.find("td", attrs={re.compile("total-rating")}).get_text(strip=True)
            logger.info(f"{txt}: {rating, yield_, dsi, strategy}")
            return [rating, yield_, dsi, strategy]
        else:
            pass
    return ['err', 'err', 'err', 'err']


def send_dohod(index_cell: int, lst: list):
    # Пример записи в файл

    values = service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": [
                {"range": f"AL{index_cell}:AO{index_cell}",
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
    read_ticker()
    get_html_code()
    clear_area()

    with open(".\data_pkl\Ticker_C_collum.pkl", 'rb') as f:
        lst = pickle.load(f)
    cell = 3
    for i in lst["values"][0]:
        if len(i) > 1:
            lst = read_from_file(i)
            send_dohod(cell, lst)
            cell += 1
        else:
            cell += 1
        time.sleep(1)

if __name__ == '__main__':
    main()
