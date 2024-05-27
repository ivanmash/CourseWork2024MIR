from plistlib import InvalidFileException
from selenium.common import NoSuchElementException, ElementClickInterceptedException, ElementNotInteractableException, \
    TimeoutException
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from bs4 import BeautifulSoup
import pandas as pd
import os
from time import sleep


def handle_exceptions(e):
    if isinstance(e, (
            NoSuchElementException, ElementClickInterceptedException, ElementNotInteractableException,
            TimeoutException)):
        print(f"Ошибка: {e.__class__.__name__}! Проблема с обработкой сайта!")
    elif isinstance(e, FileNotFoundError):
        print(f"Ошибка {e.__class__.__name__}: Файл не найден!")
    elif isinstance(e, PermissionError):
        print(f"Ошибка {e.__class__.__name__}: Отказано в доступе!")
    elif isinstance(e, InvalidFileException):
        print(f"Ошибка {e.__class__.__name__}: Недопустимый файл!")
    else:
        print(f"Произошла неизвестная ошибка: {e.__class__.__name__}!")


browser = None
try:
    browser = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
    url = 'https://tabiturient.ru/globalrating/'
    browser.get(url)
    sleep(1)
    soup = BeautifulSoup(browser.page_source, 'lxml')
    table = soup.find('div', class_='obramtop100')
    rows = table.find_all('table', class_='listtop100')
    info = []
    for row in rows:
        data = row.find_all('td')[2]
        indicators = row.find_all('td')[4]
        rating = data.find('span', class_='font2 yesmobile1').text.strip()[1:]
        name = data.find('b').text.strip()
        category = indicators.find_all('b')[0].text.strip()
        score = indicators.find_all('b')[1].text.strip()
        info.append([rating, name, score, category])

    column_names = ['Rating', 'Name', 'Grade', 'Category']
    df = pd.DataFrame(info, columns=column_names)
    print(df.info)

    data_folder = os.path.join(os.path.dirname(__file__), '..', 'Data')
    if not os.path.exists(data_folder):
        os.makedirs(data_folder)
    russia_folder = os.path.join(data_folder, 'Russia')
    if not os.path.exists(russia_folder):
        os.makedirs(russia_folder)
    file_path = os.path.join(russia_folder, 'Tabiturient.xlsx')
    df.to_excel(file_path, index=False)

except Exception as e:
    handle_exceptions(e)
finally:
    if browser is not None:
        browser.close()
