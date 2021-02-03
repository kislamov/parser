import os
import errno
import pickle
from time import sleep, time
from random import randint
from sys import platform
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.remote.webelement import WebElement
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


def rnd():
    return randint(1, 10) * 0.01


options = Options()

options.add_argument("--disable-notifications")

options.add_argument("--disable-infobars")

options.add_argument("--mute-audio")

platform_ = platform

if platform_ in ['linux', 'darwin']:
    driver = webdriver.Chrome(executable_path="./chromedriver", options=options)
else:
    driver = webdriver.Chrome(executable_path="./chromedriver.exe", options=options)

# try:
#     cookies = pickle.load(open("cookies.pkl", "rb"))
#     for cookie in cookies:
#         if cookie.get('domain'):
#             del cookie['domain']
#         print(cookie)
#         driver.add_cookie(cookie)
# except FileNotFoundError:
#     pass


driver.get("https://weather.com/ru-RU")


def get_button(key, by_=By.ID) -> WebElement:
    sleep(randint(3, 5) + rnd())
    return WebDriverWait(driver, 3).until(EC.presence_of_element_located((by_, key)))


def parse_station(station_name='карасайский'):
    day = date.today().strftime("%d_%m_%y")

    filename = f'analysis_{day}/{stations[station_name]}_weathercom.html'

    sleep(3 + rnd())

    forecast_btn = get_button('a.styles--listItem--2JY_Z:nth-child(2) > span:nth-child(1)', By.CSS_SELECTOR)

    sleep(1 + rnd())

    forecast_btn.click()

    sleep(5 + rnd())

    query_btn = get_button('LocationSearch_input', By.ID)

    sleep(rnd())

    query_btn.click()

    for i in station_name:
        query_btn.send_keys(i)
        sleep(rnd())

    sleep(2 + rnd())
    query_btn.send_keys(Keys.ENTER)

    sleep(3 + rnd())

    save_html(filename=filename, html_to_save=driver.page_source)


def save_html(filename, html_to_save):
    print(f'saving {filename}')
    if not os.path.exists(os.path.dirname(filename)):
        try:
            os.makedirs(os.path.dirname(filename))
        except OSError as exc:  # Guard against race condition
            if exc.errno != errno.EEXIST:
                raise

    with open(filename, 'w', encoding="utf-8") as f:
        f.write(html_to_save)


sleep(4 + rnd())

sleep(4 + rnd())

start = time()
for station in stations:
    parse_station(station)
end = time()

print('running time is', end - start)
pickle.dump(driver.get_cookies(), open("cookies.pkl", "wb"))
