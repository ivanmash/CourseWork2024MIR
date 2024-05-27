from selenium.webdriver.chrome.service import Service as ChromeService
from plistlib import InvalidFileException
from selenium.common import NoSuchElementException, ElementClickInterceptedException, ElementNotInteractableException, \
    TimeoutException
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

    for year in [2023, 2022, 2021, 2020, 2019]:
        url = f'https://www.scimagoir.com/rankings.php?sector=Higher+educ.&country=Africa&year={year - 6}'
        browser.get(url)
        soup = BeautifulSoup(browser.page_source, 'lxml')
        info = []
        table = soup.find('table', id='tbldata')
        rows = table.find('tbody').find_all('tr')
        for row in rows:
            if len(row.find_all('td')) == 0:
                continue
            rank = row.find_all('td')[1].text.strip().split()[0]
            global_rank = row.find_all('td')[1].text.strip().split()[1].replace('(', '').replace(')', '')
            institution = row.find_all('td')[2].text.strip()
            country = row.find_all('td')[3].text.strip()
            info.append([rank, global_rank, institution, country])
        column_names = ['Rank', 'Global Rank', 'Institution', 'Country']
        df = pd.DataFrame(info, columns=column_names)
        print(df.info)
        data_folder = os.path.join(os.path.dirname(__file__), '..', 'Data')
        if not os.path.exists(data_folder):
            os.makedirs(data_folder)
        arabic_folder = os.path.join(data_folder, 'Africa')
        if not os.path.exists(arabic_folder):
            os.makedirs(arabic_folder)
        file_path = os.path.join(arabic_folder, f'SCIMAGO_Institutions_Rankings_{year}.xlsx')
        df.to_excel(file_path, index=False)

except Exception as e:
    handle_exceptions(e)
finally:
    if browser is not None:  # Проверяем, определена ли переменная browser перед её закрытием
        browser.close()
