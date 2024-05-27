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


    # Кликаем на кнопку SCORES
    def click_scores():
        element = browser.find_element(By.ID, "scores")
        browser.execute_script("arguments[0].click();", element)


    # Кликаем на кнопку RANKING
    def click_ranking():
        element = browser.find_element(By.ID, "stats")
        browser.execute_script("arguments[0].click();", element)


    for year in [2023, 2022, 2021, 2020, 2019]:
        url = 'https://www.timeshighereducation.com/world-university-rankings/' + str(year) + '/world-ranking'
        browser.get(url)
        soup = BeautifulSoup(browser.page_source, 'lxml')
        navigation = soup.find('ul', class_='pagination')
        max_pages = int(navigation.find_all('li')[-2].text.strip())
        click_ranking()
        soup = BeautifulSoup(browser.page_source, 'lxml')
        table = soup.find('table', id='datatable-1')
        head_1 = table.find('thead').find_all('th')
        names_1 = [x.text.strip() for x in head_1]
        click_scores()
        soup = BeautifulSoup(browser.page_source, 'lxml')
        table = soup.find('table', id='datatable-1')
        head_2 = table.find('thead').find_all('th')
        names_2 = [x.text.strip() for x in head_2]
        names_2 = [x for x in names_2 if x not in names_1]
        columns = names_1 + names_2
        info = []
        for page in range(max_pages):
            click_ranking()
            soup = BeautifulSoup(browser.page_source, 'lxml')
            table_ranking = soup.find('table', id='datatable-1')
            click_scores()
            soup = BeautifulSoup(browser.page_source, 'lxml')
            table_scores = soup.find('table', id='datatable-1')
            rows_ranking = table_ranking.find('tbody').find_all('tr', role='row')
            rows_scores = table_scores.find('tbody').find_all('tr', role='row')
            for row_ranking, row_scores in zip(rows_ranking, rows_scores):
                rank = row_ranking.find_all('td')[0].text.strip()
                if rank[0] == '=':
                    rank = rank[1:]
                name_and_country = ''
                if row_ranking.find_all('td')[1].find('a', class_='ranking-institution-title') and \
                        row_ranking.find_all('td')[1].find('div', class_='location'):
                    name_and_country = row_ranking.find_all('td')[1].find('a',
                                                                          class_='ranking-institution-title').text.strip() + ' / ' + \
                                       row_ranking.find_all('td')[1].find('div', class_='location').text.strip()
                other_info_1 = [x.text.strip() if x else '' for x in row_ranking.find_all('td')[2:]]
                other_info_2 = [x.text.strip() if x else '' for x in row_scores.find_all('td')[2:]]
                current_info = [rank, name_and_country] + other_info_1 + other_info_2
                info.append(current_info)
            if page != max_pages - 1:
                while True:
                    try:
                        next_button = browser.find_element(By.CSS_SELECTOR,
                                                           'li.paginate_button.next.mz-page-processed a')
                        next_button.click()
                        break
                    except:
                        sleep(0.5)  # Пауза в пол секунды и повторяем нажатие
        df = pd.DataFrame(info, columns=columns)
        print(df.info)
        data_folder = os.path.join(os.path.dirname(__file__), '..', 'Data')
        if not os.path.exists(data_folder):
            os.makedirs(data_folder)
        arabic_folder = os.path.join(data_folder, 'World')
        if not os.path.exists(arabic_folder):
            os.makedirs(arabic_folder)
        file_path = os.path.join(arabic_folder, f'The_Higher_Education_{year}.xlsx')
        df.to_excel(file_path, index=False)

except Exception as e:
    handle_exceptions(e)
finally:
    if browser is not None:
        browser.close()
