import requests
from bs4 import BeautifulSoup
import openpyxl

URL = 'http://pulser.kz/?grpsub=106423'
HEADERS = { 'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:84.0) Gecko/20100101 Firefox/84.0',
            'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'
            }
HOST = 'http://pulser.kz/'


def get_html(url, params=None):
    r = requests.get(url, headers=HEADERS, params=params)
    return r


def get_content(html):
    soup = BeautifulSoup(html, 'html.parser')
    items = soup.find_all('div', class_='productBlock')

    pulser = [ ]
    for item in items:
        pulser.append({
            'title': item.find('span', class_='brandTop').get_text(),
            'description': item.find('h2').get_text(strip=True).replace('\xad', ''),
            'price': item.find('dd', class_='dark').find_next('dd').get_text().replace('\n', ''),
            'link': HOST + item.find('a').get('href')
        })
    del pulser[ 0 ]
    return pulser


def parse():
    html = get_html(URL)
    if html.status_code == 200:
        pulser = get_content(html.text)
        to_xlsx(pulser)
    else:
        print('Error!')


def to_xlsx(pulser):
    book = openpyxl.Workbook()
    sheet = book.active
    sheet['A1'] = 'NAME'
    sheet['B1'] = 'DESCRIPTION'
    sheet['C1'] = 'PRICE'
    sheet['D1'] = 'LINK'
    row = 2
    for i in pulser:
        sheet[row][0].value = i['title']
        sheet[row][1].value = i['description']
        sheet[row][2].value = i['price']
        sheet[row][3].value = i['link']
        row += 1
    book.save('pc_list.xlsx')


parse()
