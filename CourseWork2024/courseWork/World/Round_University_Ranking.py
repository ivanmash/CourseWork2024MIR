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


    def switch_year(year):
        switch_button = browser.find_element(By.XPATH, f'//*[@id="years"]/option[{14 - (2023 - year)}]')
        switch_button.click()


    url = 'https://roundranking.com/ranking/world-university-rankings.html#world-2023'
    browser.get(url)
    for year in [2023, 2022, 2021, 2020, 2019]:
        switch_year(year)
        sleep(2)
        soup = BeautifulSoup(browser.page_source, 'lxml')
        info = []
        table = soup.find('table', class_='big-table table-sortable tablesorter')
        rows = table.find('tbody').find_all('tr')
        for row in rows:
            rank = row.find_all('td')[0].text.strip()
            university = row.find_all('td')[1].text.strip()
            score = row.find_all('td')[2].text.strip()
            country = row.find_all('td')[3].text.strip()
            league = row.find_all('td')[5].text.strip()
            info.append([rank, university, score, country, league])
        column_names = ['Rank', 'University', 'Score', 'Country', 'League']
        df = pd.DataFrame(info, columns=column_names)
        print(df.info)
        data_folder = os.path.join(os.path.dirname(__file__), '..', 'Data')
        if not os.path.exists(data_folder):
            os.makedirs(data_folder)
        arabic_folder = os.path.join(data_folder, 'World')
        if not os.path.exists(arabic_folder):
            os.makedirs(arabic_folder)
        file_path = os.path.join(arabic_folder, f'Round_University_Ranking_{year}.xlsx')
        df.to_excel(file_path, index=False)

except Exception as e:
    handle_exceptions(e)
finally:
    if browser is not None:
        browser.close()
