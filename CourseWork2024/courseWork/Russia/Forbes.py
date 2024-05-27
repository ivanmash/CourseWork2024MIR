from plistlib import InvalidFileException
from selenium.common import NoSuchElementException, ElementClickInterceptedException, ElementNotInteractableException, \
    TimeoutException
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
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
    url = 'https://datawrapper.dwcdn.net/en9ux/1/'
    browser.get(url)

    soup = BeautifulSoup(browser.page_source, 'lxml')

    info = []
    for i in range(10):
        soup = BeautifulSoup(browser.page_source, 'lxml')
        table = soup.find('table').find('tbody')
        rows = table.find_all('tr')
        for j in range(len(rows)):
            row = rows[j]
            name = row.find('th').text.strip()
            other_params = [x.text.strip() for x in row.find_all('td')]
            info.append([name, *other_params])
        if i != 9:
            sleep(1)
            next_page_button = browser.find_element(By.CLASS_NAME, 'next.svelte-1ya2siw')
            sleep(1)
            next_page_button.click()
            sleep(1)

    column_names = ['Название', 'Востребованность выпускников', 'Качество нетворкинга',
                    'Международная репутация', 'Качество преподавания', 'Фактор Forbes', 'Итог']
    df = pd.DataFrame(info, columns=column_names)
    print(df.info)

    # Путь к папке Data
    data_folder = os.path.join(os.path.dirname(__file__), '..', 'Data')

    # Проверка существования папки Data и ее создание, если она не существует
    if not os.path.exists(data_folder):
        os.makedirs(data_folder)

    # Путь к папке Russia внутри папки Data
    russia_folder = os.path.join(data_folder, 'Russia')

    # Проверка существования папки Russia внутри папки Data и ее создание, если она не существует
    if not os.path.exists(russia_folder):
        os.makedirs(russia_folder)

    # Путь к файлу Forbes.xlsx в папке Russia
    file_path = os.path.join(russia_folder, 'Forbes.xlsx')

    # Сохранение DataFrame в файл
    df.to_excel(file_path, index=False)

except Exception as e:
    handle_exceptions(e)
finally:
    if browser is not None:
        browser.close()

