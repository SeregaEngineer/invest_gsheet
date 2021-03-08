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


def read_ticker(range_read_ticker:str):
    """
    read ticket from sheet
    """
    # Read ticket for RBK, column -D-
    values = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range= range_read_ticker,
        majorDimension='COLUMNS'
    ).execute()
    with open(".\data_pkl\Ticker_D_collum_RBK.pkl", 'wb') as f:
        pickle.dump(values, f)


def clear_area(range: str):
    """
    Clear cell in area
    """
    service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range= range,
    ).execute()


def get_date(tiker: str):
    """

    :param tiker: number RBK Url
    :return: If server Ok = True
    """
    url = f'https://quote.rbc.ru/ticker/{tiker}'
    headers = {'accept': '*/*',
               'User-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:67.0) Gecko/20100101 Firefox/67.0'}

    r = requests.get(url=url, headers=headers)

    if r.status_code == 200:
        soup = bs(r.text, 'lxml')
        with open('.\data_pkl\data_from_RBK.pkl', 'wb') as file:
            pickle.dump(str(soup), file)
        return True
    else:
        logger.error(f"server error: {r.status_code}")
        return False


def parser():
    """
    :return:list actual + forecast + dividend
    """
    forec = []
    # load
    with open('.\data_pkl\data_from_RBK.pkl', 'rb') as f:
        soup = bs(pickle.load(f), 'lxml')

    price = soup.find('div', attrs={'class', 'chart__info'})

    # Получаем Название, актульную цену и изменени в цене
    name_issuer = price.find('span', attrs={'span', 'chart__info__name-short'}).get_text(strip=True)
    price_actual = price.find('span', attrs={'class', 'chart__info__sum'}).get_text(strip=True)
    price_change = price.find('span', attrs={'class', 'chart__info__change chart__change'}).get_text(strip=True)[1:-1]
    actual_data = [name_issuer, price_actual, price_change]

    forec.append(actual_data)

    # Иcтория дивидендов
    divid = soup.find_all('span', attrs={'class', 'q-item__dividend__item'})
    forec.append(dividend(divid))

    # Получаем прогноз (forecast)
    at = soup.find('div', attrs={'class', 'js-review'})
    a1_obj = at.find_all('span', attrs={'class', 'q-item__review__container'})
    for i in a1_obj:
        forec.append(forecast(i))

    if len(forec) < 7:
        x = 7 - len(forec)
        for i in range(x):
            lst = ['-', '-', '-', '-']
            forec.append(lst)

    logger.info(forec)

    return forec


def forecast(var_object: object) -> object:
    """
    Return value forecast (Прогноз)
    :param var_object:  bs4 object
    :return: list [name analyst, analytics date, forecast, reliability ]
    """
    obj = var_object.select(".q-item__review__inner")
    x = 0
    for i in obj[:5]:
        if x == 0:
            try:
                forecast = i.find('span', attrs=('class', 'q-item__review__sum')).get_text(strip=True)
            except:
                forecast = "Err"
        elif x == 1:
            try:
                date = i.find('span', attrs=('class', 'q-item__review__value')).get_text(strip=True)
            except:
                date = "Err"
        elif x == 2:
            try:
                analyst = i.find('span', attrs=('class', 'q-item__review__value')).get_text(strip=True)
            except:
                analyst = "Err"
        elif x == 3:
            try:
                realib = i.find('span', attrs=('class', 'q-item__review__value')).get_text(strip=True)
            except:
                realib = "Err"
        x += 1
    # logger.info("forecast = OK")

    return [analyst, date, forecast, realib]


def dividend(var_object: object) -> object:
    """
    Return statistic history dividend
    :param var_object:  bs4 object
    :return: list [[data,count,percent],[***],[***] ]
    """
    history_div = []

    for i in var_object[:3]:
        lst = []
        try:
            data = i.find('span', attrs={'class', 'q-item__dividend__inner'}).get_text(strip=True)
        except:
            data = "Err"
        try:
            count = i.find('span', attrs={'class', 'q-item__dividend__sum'}).get_text(strip=True)
        except:
            count = "Err"
        try:
            percent = i.find('span', attrs={'class', 'q-item__dividend__percent'}).get_text(strip=True)
        except:
            percent = "Err"

        lst = [data, count, percent]
        history_div.append(lst)
    # logger.info("dividend = OK")
    if len(history_div) < 3:
        for i in range(3 - len(history_div)):
            lst = ['-', '-', '-']
            history_div.append(lst)

    return history_div


def send_rbk(index_cell: int, lst: list):
    # Пример записи в файл

    values = service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": [
                {"range": f"F{index_cell}:AK{index_cell}",
                 "majorDimension": "ROWS",
                 "values": [
                     [  # f'{date.strftime("%d.%m.%Y")}'                            #Дата обновления
                         f'{lst[0][0]}', f'{lst[0][1]}', f'{lst[0][2]}',  # Актуальная инфа
                         f'{lst[1][0][0]}', f'{lst[1][0][1]}', f'{lst[1][0][2]}',  # Дивиденды 1
                         f'{lst[1][1][0]}', f'{lst[1][1][1]}', f'{lst[1][1][2]}',  # Дивиденды 2
                         f'{lst[1][2][0]}', f'{lst[1][2][1]}', f'{lst[1][2][2]}',  # Дивиденды 3
                         f'{lst[2][0]}', f'{lst[2][1]}', f'{lst[2][2]}', f'{lst[2][3]}',  # Аналитик 1
                         f'{lst[3][0]}', f'{lst[3][1]}', f'{lst[3][2]}', f'{lst[3][3]}',  # Аналитик 2
                         f'{lst[4][0]}', f'{lst[4][1]}', f'{lst[4][2]}', f'{lst[4][3]}',  # Аналитик 3
                         f'{lst[5][0]}', f'{lst[5][1]}', f'{lst[5][2]}', f'{lst[5][3]}',  # Аналитик 4
                         f'{lst[6][0]}', f'{lst[6][1]}', f'{lst[6][2]}', f'{lst[6][3]}',  # Аналитик 5

                     ],
                 ]
                 },
            ]
        }
    ).execute()


def main(range_read_ticker: str):

    read_ticker(range_read_ticker)

    with open(".\data_pkl\Ticker_D_collum_RBK.pkl", 'rb') as f:
        val = pickle.load(f)
    cell = 3
    for i in range(len(val['values'][0])):
        if (val['values'][0][i] is not None) and get_date(val['values'][0][i]):
            lst = parser()
            send_rbk(cell, lst)
            cell += 1
        else:
            cell += 1
            pass
        time.sleep(1)


if __name__ == '__main__':
    main()

