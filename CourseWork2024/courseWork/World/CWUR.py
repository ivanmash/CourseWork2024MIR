from plistlib import InvalidFileException
from selenium.common import NoSuchElementException, ElementClickInterceptedException, ElementNotInteractableException, \
    TimeoutException
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from bs4 import BeautifulSoup
import pandas as pd
import os


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

    url = 'https://cwur.org/2023.php'
    browser.get(url)
    soup = BeautifulSoup(browser.page_source, 'lxml')

    info = []
    table = soup.find('table', id='cwurTable')
    rows = table.find('tbody').find_all('tr')
    column_names = table.find('thead').text.strip().split('\n')
    for row in rows:
        info.append([x.text.strip().replace('\xa0', ' ').replace('\n', ';') for x in row.find_all('td')])

    df = pd.DataFrame(info, columns=column_names)
    print(df.info)

    data_folder = os.path.join(os.path.dirname(__file__), '..', 'Data')
    if not os.path.exists(data_folder):
        os.makedirs(data_folder)
    arabic_folder = os.path.join(data_folder, 'World')
    if not os.path.exists(arabic_folder):
        os.makedirs(arabic_folder)
    file_path = os.path.join(arabic_folder, 'CWUR.xlsx')
    df.to_excel(file_path, index=False)

except Exception as e:
    handle_exceptions(e)
finally:
    if browser is not None:
        browser.close()
