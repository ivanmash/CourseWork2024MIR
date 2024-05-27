from plistlib import InvalidFileException
from selenium.common import NoSuchElementException, ElementClickInterceptedException, ElementNotInteractableException, \
    TimeoutException
import requests
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
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


    def switch_page(arg):
        suffixes = {'alumni': '[1]', 'award': '[2]', 'hiCi': '[3]', 'n_and_s': '[4]', 'pub': '[5]', 'pcp': '[6]'}
        switch_page = browser.find_element(By.XPATH,
                                           '//*[@id="content-box"]/div[2]/table/thead/tr/th[6]/div/div[1]/div[1]')
        switch_page.click()
        new_page = browser.find_element(By.XPATH,
                                        '//*[@id="content-box"]/div[2]/table/thead/tr/th[6]/div/div[1]/div[2]/ul/li'
                                        + suffixes[arg])
        new_page.click()


    def add_extra_info(intermediate_info):
        extra_soup = BeautifulSoup(browser.page_source, 'lxml')
        extra_table = extra_soup.find('table', class_='rk-table')
        extra_table_rows = extra_table.find_all('tr', {'data-v-ae1ab4a8': True})
        for j in range(1, len(extra_table_rows)):
            extra_row = extra_table_rows[j]
            extra_data = extra_row.find_all('td')
            extra_info = extra_data[5].text.strip()
            intermediate_info[j - 1].append(extra_info)
        return intermediate_info


    def switch_region(region, year) -> bool:
        number = {2023: {'Egypt': 18, 'Pakistan': 43, 'Iran': 29, 'Qatar': 46, 'Saudi Arabia': 48, 'UAE': 61},
                  2022: {'Egypt': 18, 'Pakistan': 44, 'Iran': 28, 'Qatar': 47, 'Saudi Arabia': 50, 'UAE': 62},
                  2021: {'Egypt': 18, 'Pakistan': 43, 'Iran': 28, 'Qatar': 46, 'Saudi Arabia': 49, 'UAE': -1},
                  2020: {'Egypt': 19, 'Pakistan': 42, 'Iran': 29, 'Qatar': 45, 'Saudi Arabia': 48, 'UAE': -1},
                  2019: {'Egypt': 19, 'Pakistan': 43, 'Iran': 28, 'Qatar': -1, 'Saudi Arabia': 48, 'UAE': 61}}
        switch_page = browser.find_element(By.XPATH,
                                           '//*[@id="content-box"]/div[2]/table/thead/tr/th[3]/div/div/div[1]/input')
        number = number[year][region]
        if number == -1:
            return False
        switch_page.click()
        new_page = browser.find_element(By.XPATH,
                                        f'//*[@id="content-box"]/div[2]/table/thead/tr/th[3]/div/div/div[2]/ul/li[{number}]')

        new_page.click()
        return True


    for year in [2023, 2022, 2021, 2020, 2019]:
        url = 'https://www.shanghairanking.com/rankings/arwu/' + str(year)
        browser.get(url)
        overall_information = []
        for region in ['Egypt', 'Pakistan', 'Iran', 'Qatar', 'Saudi Arabia', 'UAE']:
            if not (switch_region(region, year)):
                continue
            soup = BeautifulSoup(browser.page_source, 'lxml')
            max_pages = 1
            try:
                pages = soup.find('ul', class_='ant-pagination').find_all('li')
                max_pages = int(pages[-2].text.strip())
            except Exception:
                pass
            for page in range(max_pages):
                # Скрипт для прокрутки страницы вниз
                scroll_script = "window.scrollTo(0, document.body.scrollHeight);"
                browser.execute_script(scroll_script)
                try:
                    next_page = browser.find_element(By.XPATH, f'//li[@title="{str(page + 1)}"]')
                    next_page.click()
                except Exception:
                    pass
                soup = BeautifulSoup(browser.page_source, 'lxml')
                table = soup.find('table', class_='rk-table')
                table_rows = table.find_all('tr', {'data-v-ae1ab4a8': True})
                intermediate_info = []
                for i in range(1, len(table_rows)):
                    row = table_rows[i]
                    data = row.find_all('td')
                    world_rank = data[0].text.strip()
                    institution = data[1].find('span').text.strip()
                    national_rating = data[3].text.strip()
                    total_score = data[4].text.strip()
                    alumni = data[5].text.strip()
                    link = 'https://www.shanghairanking.com' + data[1].find('a').get('href')
                    info_response = requests.get(link)
                    info_page = BeautifulSoup(info_response.content, 'html.parser')
                    country = \
                        info_page.find('div', class_='contact-msg').find_all('div', class_='contact-item')[1].find_all(
                            'span')[
                            1].text.strip()
                    intermediate_info.append([world_rank, institution, country, national_rating, total_score, alumni])
                # Скрипт для прокрутки страницы вверх
                scroll_script = "window.scrollTo(0, 0);"
                browser.execute_script(scroll_script)
                switch_page('award')
                intermediate_info = add_extra_info(intermediate_info)
                switch_page('hiCi')
                intermediate_info = add_extra_info(intermediate_info)
                switch_page('n_and_s')
                intermediate_info = add_extra_info(intermediate_info)
                switch_page('pub')
                intermediate_info = add_extra_info(intermediate_info)
                switch_page('pcp')
                intermediate_info = add_extra_info(intermediate_info)
                switch_page('alumni')
                for elem in intermediate_info:
                    overall_information.append(elem)
        column_names = ['World Rank', 'Institution', 'Country', 'National Rank', 'Total Score', 'Alumni', 'Award',
                        'HiCi',
                        'N&S', 'PUB', 'PCP']
        df = pd.DataFrame(overall_information, columns=column_names)
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

        file_path = os.path.join(arabic_folder, f'Shanghai_Ranking_{str(year)}.xlsx')

        # Сохранение DataFrame в файл
        df.to_excel(file_path, index=False)

except Exception as e:
    handle_exceptions(e)
finally:
    if browser is not None:
        browser.close()
