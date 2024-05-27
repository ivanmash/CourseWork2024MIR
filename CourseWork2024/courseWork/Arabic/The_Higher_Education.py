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

    for year in [2023, 2022, 2021]:
        url = f'https://www.timeshighereducation.com/world-university-rankings/{year}/arab-university-rankings'
        browser.get(url)
        soup = BeautifulSoup(browser.page_source, 'lxml')
        navigation = soup.find('ul', class_='pagination')
        max_pages = int(navigation.find_all('li')[-2].text.strip())
        table = soup.find('table', id='datatable-1')
        head = table.find('thead').find_all('th')
        column_names = [x.text.strip() for x in head]
        info = []
        for page in range(max_pages):
            soup = BeautifulSoup(browser.page_source, 'lxml')
            table = soup.find('table', id='datatable-1')
            rows = table.find('tbody').find_all('tr', role='row')
            for row in rows:
                rank = row.find_all('td')[0].text.strip()
                if rank[0] == '=':
                    rank = rank[1:]
                name_and_country = ''
                if row.find_all('td')[1].find('a', class_='ranking-institution-title') and \
                        row.find_all('td')[1].find('div', class_='location'):
                    name_and_country = row.find_all('td')[1].find('a',
                                                                  class_='ranking-institution-title').text.strip() + ' / ' + \
                                       row.find_all('td')[1].find('div', class_='location').text.strip()
                other_info = [x.text.strip() if x else '' for x in row.find_all('td')[2:]]
                current_info = [rank, name_and_country] + other_info
                info.append(current_info)
            if page != max_pages - 1:
                while True:
                    try:
                        next_button = browser.find_element(By.CSS_SELECTOR, 'li.paginate_button.next.mz-page-processed a')
                        next_button.click()
                        break
                    except:
                        sleep(0.5)  # Пауза в пол секунды и повторяем нажатие
        df = pd.DataFrame(info, columns=column_names)
        print(df.info)
        data_folder = os.path.join(os.path.dirname(__file__), '..', 'Data')

        # Проверка существования папки Data и ее создание, если она не существует
        if not os.path.exists(data_folder):
            os.makedirs(data_folder)

        # Путь к папке Russia внутри папки Data
        arabic_folder = os.path.join(data_folder, 'Arabic')

        # Проверка существования папки Arabic внутри папки Data и ее создание, если она не существует
        if not os.path.exists(arabic_folder):
            os.makedirs(arabic_folder)

        file_path = os.path.join(arabic_folder, f'The_Higher_Education_{str(year)}.xlsx')

        # Сохранение DataFrame в файл
        df.to_excel(file_path, index=False)

except Exception as e:
    handle_exceptions(e)
finally:
    if browser is not None:
        browser.close()
