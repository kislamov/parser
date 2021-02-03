import requests
from bs4 import BeautifulSoup
import openpyxl
from datetime import datetime

URL = 'https://weather.com/ru-RU/weather/hourbyhour/l/91fc832467e02603d03bbbb6082d330bcd1f1cf7997e3354cfc84cb3446443f8'
HEADERS = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:84.0) Gecko/20100101 Firefox/84.0',
           'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'
           }

stations = {
    'карасайский': 'karasaysky',
    'баиркум': 'baiyrkum',
    'державинск': 'derzhavinsk',
    'токмансай': 'tokmansai',
    'илийский': 'iliysky',
    'егиндыбулак, карагандинская': 'egindybulak',
    'аккудук': 'akkuduk',
    'бестобе, акмолинская': 'bestobe',
    'ерейментау': 'ereimentau',
    'степное': 'stepnoe',
    'жансугуров': 'zhansugurov',
    'коныролен': 'konyrolen',
    'исатай': 'isatai',
    'макат': 'makat',
}


def get_html(url, params=None):
    r = requests.get(url, headers=HEADERS, params=params)
    return r


def get_location(html):
    soup = BeautifulSoup(html, 'html.parser')
    location = soup.find('span', class_='LocationPageTitle--PresentationName--Injxu').text
    return location


def get_content(html):
    soup = BeautifulSoup(html, 'html.parser')
    items = soup.find_all('div', class_='DaypartDetails--DetailSummaryContent--1c28m Disclosure--SummaryDefault--1z_mF')
    # print(items)
    forecast = []
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
        for j in i['wind']:
            if j in ' СЮЗВ':
                i['wind'] = i['wind'].replace(j, '')

        i['wind'] = round(int(i['wind']) * 1000 / 3600, 2)

    # print(forecast)
    # print(len(forecast))
    return forecast


def parse():
    html = get_html(URL)
    if html.status_code == 200:
        forecast = get_content(html.text)
        location = get_location(html.text).split(',')
        if location[0].lower() in stations:
            to_xlsx(forecast, stations[location[0].lower()])
    else:
        print('Error!')


def to_xlsx(forecast, location):
    book = openpyxl.Workbook()
    sheet = book.active
    sheet['A1'] = 'Время, часы'
    sheet['B1'] = 'Температура, °C'
    sheet['C1'] = 'Вероятность осадков, %'
    sheet['D1'] = 'Скорость ветра, м/c'
    sheet['E1'] = 'Состояние погоды'
    row = 2
    for i in forecast:
        sheet[row][0].value = i['hour']
        sheet[row][1].value = i['temp']
        sheet[row][2].value = i['probability']
        sheet[row][3].value = i['wind']
        sheet[row][4].value = i['condition']
        row += 1
    book.save(f'{location}.xlsx')
    book.close()


parse()
