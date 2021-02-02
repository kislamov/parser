import requests
from bs4 import BeautifulSoup
import openpyxl

URL = 'https://weather.com/ru-RU/weather/hourbyhour/l/78eac41c89ce940fa938848cf6deee49e840d217a3c2427ceaaae24d4888b9dbbc0ac17bdebe7c6db341e77385fb90e0'
HEADERS = { 'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:84.0) Gecko/20100101 Firefox/84.0',
            'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'
            }
HOST = 'http://pulser.kz/'


def get_html(url, params=None):
    r = requests.get(url, headers=HEADERS, params=params)
    return r


def get_content(html):
    soup = BeautifulSoup(html, 'html.parser')
    items = soup.find_all('div', class_='DaypartDetails--DetailSummaryContent--1c28m Disclosure--SummaryDefault--1z_mF')
    # print(items)
    forecast = []
    for item in items:
        forecast.append({
            'hour': item.find('h2', class_='DetailsSummary--daypartName--1Mebr').text,
            'temp': item.find('span', class_='DetailsSummary--tempValue--RcZzi').text,
            'condition': item.find('span', class_='DetailsSummary--extendedData--aaFeV').text,
            'probability': item.find('div', class_='DetailsSummary--precip--2ARnx').find_next('span').text,
            'wind': item.find('span', class_='Wind--windWrapper--1Va1P undefined').text
        })
    # print(forecast)
    # print(len(forecast))
    return forecast


def parse():
    html = get_html(URL)
    if html.status_code == 200:
        forecast = get_content(html.text)
        to_xlsx(forecast)

    else:
        print('Error!')


def to_xlsx(forecast):
    book = openpyxl.Workbook()
    sheet = book.active
    sheet['A1'] = 'hour'
    sheet['B1'] = 'temp'
    sheet['C1'] = 'condition'
    sheet['D1'] = 'probability'
    row = 2
    for i in forecast:
        sheet[row][0].value = i['hour']
        sheet[row][1].value = i['temp']
        sheet[row][2].value = i['condition']
        sheet[row][3].value = i['probability']
        row += 1
    book.save('forecasting.xlsx')


parse()
