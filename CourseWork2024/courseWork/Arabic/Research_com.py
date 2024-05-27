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

    url = 'https://research.com/university-rankings/best-global-universities/sa'
    browser.get(url)
    soup = BeautifulSoup(browser.page_source, 'lxml')
    table = soup.find('div', id='rankingItems')
    rows = table.find_all('div', class_='cols university-item rankings-content__item show')
    info = []
    for row in rows:
        global_rank = row.find('span', class_='col col--3 py-0 px-0 position border').text.strip().split('\n')[0]
        national_rank = row.find('span', class_='col col--3 py-0 px-0 position').text.strip().split('\n')[0]
        name = row.find('h4').text.strip()
        country = row.find('span', class_='sh').text.strip()
        scholars = row.find('span', class_='ranking no-wrap').text.strip()
        H_index = row.find_all('span', class_='ranking no-wrap')[1].text.strip()
        publications = row.find('span', class_='col col--3 py-0 ranking no-wrap').text.strip()
        info.append([name, country, global_rank, national_rank, scholars, publications, H_index])
    column_names = ['Name', 'Country', 'World Rank', 'National Rank', 'Scholars', 'Publications', 'H-Index']

    df = pd.DataFrame(info, columns=column_names)
    print(df.info)
    # Путь к папке Data
    data_folder = os.path.join(os.path.dirname(__file__), '..', 'Data')

    # Проверка существования папки Data и ее создание, если она не существует
    if not os.path.exists(data_folder):
        os.makedirs(data_folder)

    # Путь к папке Russia внутри папки Data
    arabic_folder = os.path.join(data_folder, 'Arabic')

    # Проверка существования папки Arabic внутри папки Data и ее создание, если она не существует
    if not os.path.exists(arabic_folder):
        os.makedirs(arabic_folder)

    # Путь к файлу Research_com.xlsx в папке Arabic
    file_path = os.path.join(arabic_folder, 'Research_com.xlsx')

    # Сохранение DataFrame в файл
    df.to_excel(file_path, index=False)

except Exception as e:
    handle_exceptions(e)
finally:
    if browser is not None:
        browser.close()
