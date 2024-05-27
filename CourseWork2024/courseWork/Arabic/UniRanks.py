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

    url = 'https://www.uniranks.com/ranking?region=Arab+World'
    browser.get(url)
    soup = BeautifulSoup(browser.page_source, 'lxml')
    rows = soup.find_all('div', 'row mb-2 uni-listing-wrapper md-hidden')
    navigation = soup.find('ul', class_='pagination')
    max_pages = int(navigation.find_all('li')[-2].text.strip())
    info = []
    for page in range(1, max_pages + 1):
        url = f'https://www.uniranks.com/ranking?region=Arab+World&page={page}'
        browser.get(url)
        soup = BeautifulSoup(browser.page_source, 'lxml')
        rows = soup.find_all('div', 'row mb-2 uni-listing-wrapper md-hidden')
        for row in rows:
            name = row.find('a').text.strip()
            other_info = row.find('div', class_='mb-6 px-2')
            current_info = [x.text.strip() if x else '' for x in other_info.find_all('span', class_='text-muted')]
            info.append([name, *current_info])
    column_names = ['Name', 'Rank', 'Score', 'Location']
    df = pd.DataFrame(info, columns=column_names)
    print(df.info)

    data_folder = os.path.join(os.path.dirname(__file__), '..', 'Data')
    if not os.path.exists(data_folder):
        os.makedirs(data_folder)
    arabic_folder = os.path.join(data_folder, 'Arabic')
    if not os.path.exists(arabic_folder):
        os.makedirs(arabic_folder)
    file_path = os.path.join(arabic_folder, 'UniRanks.xlsx')
    df.to_excel(file_path, index=False)

except Exception as e:
    handle_exceptions(e)
finally:
    if browser is not None:
        browser.close()
