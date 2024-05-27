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

    for year in [2023, 2022, 2021, 2020, 2019]:
        url = f'https://www.universityrankings.ch/en/results/QS/{year}'
        browser.get(url)
        soup = BeautifulSoup(browser.page_source, 'lxml')
        nav = soup.find_all('div', class_='navbar container right')[0]
        pages = nav.find_all('a')
        max_pages = int(pages[4].text.strip())

        info = []
        for i in range(max_pages):
            soup = BeautifulSoup(browser.page_source, 'lxml')
            table = soup.find('table', id='RankingResults').find('tbody')
            rows = table.find_all('tr', itemprop='itemListElement')
            for row in rows:
                world_rank = row.find_all('td')[0].text.strip().split()[0]
                institution = row.find_all('td')[1].text.strip()
                country = row.find_all('td')[2].find('a').get('title').split()[2]
                info.append([world_rank, institution, country])
            if i != max_pages - 1:
                sleep(1)
                next_button = browser.find_element(By.XPATH, "//a[contains(., 'Next')]")
                sleep(1)
                next_button.click()

        column_names = ['World Rank', 'Institution', 'Country']
        df = pd.DataFrame(info, columns=column_names)
        print(df.info())

        data_folder = os.path.join(os.path.dirname(__file__), '..', 'Data')
        if not os.path.exists(data_folder):
            os.makedirs(data_folder)
        arabic_folder = os.path.join(data_folder, 'World')
        if not os.path.exists(arabic_folder):
            os.makedirs(arabic_folder)
        file_path = os.path.join(arabic_folder, f'QS_Rankings_{year}.xlsx')
        df.to_excel(file_path, index=False)

except Exception as e:
    handle_exceptions(e)
finally:
    if browser is not None:
        browser.close()
