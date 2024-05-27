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


    def switch_page(arg):
        suffixes = {'alumni': '[1]', 'award': '[2]', 'hiCi': '[3]', 'n_and_s': '[4]', 'pub': '[5]', 'pcp': '[6]'}
        switch_page = browser.find_element(By.XPATH,
                                           '//*[@id="content-box"]/div[2]/table/thead/tr/th[6]/div/div[1]/div[1]')
        switch_page.click()
        new_page = browser.find_element(By.XPATH,
                                        '//*[@id="content-box"]/div[2]/table/thead/tr/th[6]/div/div[1]/div[2]/ul/li'
                                        + suffixes[arg])
        new_page.click()


    def switch_region(year):
        number = 47
        if year == 2022:
            number = 49
        elif year == 2021:
            number = 48
        switch_page = browser.find_element(By.XPATH,
                                           '//*[@id="content-box"]/div[2]/table/thead/tr/th[3]/div/div/div[1]/input')
        switch_page.click()
        new_page = browser.find_element(By.XPATH,
                                        f'//*[@id="content-box"]/div[2]/table/thead/tr/th[3]/div/div/div[2]/ul/li[{number}]')

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


    for year in [2023, 2022, 2021, 2020, 2019]:
        url = 'https://www.shanghairanking.com/rankings/arwu/' + str(year)
        browser.get(url)
        soup = BeautifulSoup(browser.page_source, 'lxml')
        info = []
        switch_region(year)
        soup = BeautifulSoup(browser.page_source, 'lxml')
        table = soup.find('table', class_='rk-table')
        table_rows = table.find_all('tr', {'data-v-ae1ab4a8': True})
        for i in range(1, len(table_rows)):
            row = table_rows[i]
            data = row.find_all('td')
            world_rank = data[0].text.strip()
            institution = data[1].find('span').text.strip()
            national_rating = data[3].text.strip()
            total_score = data[4].text.strip()
            alumni = data[5].text.strip()
            info.append([world_rank, institution, national_rating, total_score, alumni])
        for extra_param in ['award', 'hiCi', 'n_and_s', 'pub', 'pcp', 'alumni']:
            sleep(1)
            switch_page(extra_param)
            if extra_param != 'alumni':
                add_extra_info(info)
        column_names = ['World Rank', 'Institution', 'National Rank', 'Total Score', 'Alumni', 'Award', 'HiCi', 'N&S',
                        'PUB', 'PCP']
        df = pd.DataFrame(info, columns=column_names)
        print(df.info)
        data_folder = os.path.join(os.path.dirname(__file__), '..', 'Data')
        if not os.path.exists(data_folder):
            os.makedirs(data_folder)
        russia_folder = os.path.join(data_folder, 'Russia')
        if not os.path.exists(russia_folder):
            os.makedirs(russia_folder)
        file_path = os.path.join(russia_folder, f'Shanghai_Ranking_{str(year)}.xlsx')
        df.to_excel(file_path, index=False)
except Exception as e:
    handle_exceptions(e)
finally:
    if browser is not None:
        browser.close()
