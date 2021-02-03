import errno
import os

from bs4 import BeautifulSoup
import openpyxl
from datetime import date


stations = {
    'карасайский': 'karasaysky',
    'байыркум': 'baiyrkum',
    'державинск': 'derzhavinsk',
    'токмансай': 'tokmansai',
    'илийский': 'iliysky',
    'егиндыбулак, караганда': 'egindybulak',
    'бестобе, акмолинск': 'bestobe',
    'ерейментау': 'ereimentau',
    'степное, актюб': 'stepnoe',
    'жансугуров': 'zhansugurov',
    'коныролен': 'konyrolen',
    'исатай, атырау': 'isatai',
    'макат, атырау': 'makat',
}



def get_content(html):
    soup = BeautifulSoup(html, 'html.parser')
    items = soup.find_all('div', class_='DaypartDetails--DetailSummaryContent--1c28m Disclosure--SummaryDefault--1z_mF')
    # print(items)
    forecast = [ ]
    for item in items:
        forecast.append({
            'hour': item.find('h2', class_='DetailsSummary--daypartName--1Mebr').text,
            'temp': int(item.find('span', class_='DetailsSummary--tempValue--RcZzi').text.replace('°', '')),
            'probability': int(
                item.find('div', class_='DetailsSummary--precip--2ARnx').find_next('span').text.replace('%', '')),
            'wind': item.find('span', class_='Wind--windWrapper--1Va1P undefined').text.replace('км/ч', ''),
            'condition': item.find('span', class_='DetailsSummary--extendedData--aaFeV').text,
        })

    for i in forecast:
        for j in i[ 'wind' ]:
            if j in ' СЮЗВ':
                i[ 'wind' ] = i[ 'wind' ].replace(j, '')

        i[ 'wind' ] = round(int(i[ 'wind' ]) * 1000 / 3600, 2)

    # print(forecast)
    # print(len(forecast))
    return forecast


def parse():
    day = date.today().strftime("%d_%m_%y")
    for station in stations:
        with open(f'analysis_03_02_21/{stations[ station ]}_weathercom.html', 'r', encoding='utf-8') as file:
            result = get_content(file.read())
        to_xlsx(result, f'result_{day}/{stations[station]}.xlsx')


def to_xlsx(forecast, filename):
    book = openpyxl.Workbook()
    sheet = book.active
    sheet[ 'A1' ] = 'Время, часы'
    sheet[ 'B1' ] = 'Температура, °C'
    sheet[ 'C1' ] = 'Вероятность осадков, %'
    sheet[ 'D1' ] = 'Скорость ветра, м/c'
    sheet[ 'E1' ] = 'Состояние погоды'
    row = 2
    for i in forecast:
        sheet[ row ][ 0 ].value = i[ 'hour' ]
        sheet[ row ][ 1 ].value = i[ 'temp' ]
        sheet[ row ][ 2 ].value = i[ 'probability' ]
        sheet[ row ][ 3 ].value = i[ 'wind' ]
        sheet[ row ][ 4 ].value = i[ 'condition' ]
        row += 1

    book.save(f'{filename}.xlsx')
    book.close()


parse()
