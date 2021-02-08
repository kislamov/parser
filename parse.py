import errno
from datetime import date
import os, os.path
import openpyxl
from bs4 import BeautifulSoup

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

day = date.today().strftime("%d_%m_%y")

os.makedirs(f'result_{day}')


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


    # print(len(forecast))
    return forecast


def parse():
    for station in stations:
        with open(f'analysis_{day}/{stations[station]}_weathercom.html', 'r', encoding='utf-8') as file:
            result = get_content(file.read())
        to_xlsx(result, f'{stations[station]}_weathercom.xlsx')


def to_xlsx(forecast, filename):
    book = openpyxl.Workbook()
    sheet = book.active
    sheet['A1'] = f"{'/'.join(day.split('_'))} + 2-e суток"
    sheet['A2'] = 'Время, часы'
    sheet['B2'] = 'Температура, °C'
    sheet['C2'] = 'Вероятность осадков, %'
    sheet['D2'] = 'Скорость ветра, м/c'
    sheet['E2'] = 'Состояние погоды'
    row = 3
    for i in forecast:
        sheet[row][0].value = i['hour']
        sheet[row][1].value = i['temp']
        sheet[row][2].value = i['probability']
        sheet[row][3].value = i['wind']
        sheet[row][4].value = i['condition']
        row += 1
    print(f'saving {filename}')

    book.save(f'result_{day}/{filename}')
    book.close()


parse()
