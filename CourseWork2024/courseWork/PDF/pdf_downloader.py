from plistlib import InvalidFileException

from selenium.common import NoSuchElementException, ElementClickInterceptedException, ElementNotInteractableException, \
    TimeoutException
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
import pandas as pd
import requests
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


# Функция для поиска и скачивания PDF-отчетов по ключевым словам
def download_pdf_reports(driver, institute_name, report_url, keywords):
    try:
        if report_url:
            driver.get(report_url)
            pdf_links = driver.find_elements(By.TAG_NAME, 'a')
            institute_folder = os.path.join(data_folder, institute_name)

            if len(pdf_links) > 0:
                if not os.path.exists(institute_folder):
                    os.makedirs(institute_folder)

            for pdf_link in pdf_links:
                try:
                    href = pdf_link.get_attribute('href')
                    if href and href.endswith('.pdf'):
                        # Проверяем, содержит ли ссылка ключевые слова
                        if any(keyword.lower() in href.lower() for keyword in keywords):
                            file_name = os.path.join(institute_folder, href.split('/')[-1])
                            response = requests.get(href, stream=True)
                            with open(file_name, 'wb') as pdf_file:
                                pdf_file.write(response.content)
                except Exception as e:
                    continue
    except Exception as e:
        print(f"Ошибка при обработке института {institute_name}")
        print(e)


browser = None
try:
    # Чтение данных из файла links.xlsx
    current_script_path = os.path.abspath(__file__)
    current_script_directory = os.path.dirname(current_script_path)
    df = pd.read_excel(os.path.join(current_script_directory, 'links.xlsx'))
    # Удаляем строки с пустыми значениями в столбцах 1 и 3
    df = df.dropna(subset=df.columns[[1, 3]])
    data_folder = os.path.join(os.path.dirname(__file__), '..', 'PDF_reports')
    if not os.path.exists(data_folder):
        os.makedirs(data_folder)

    # Запрос ключевых слов у пользователя
    keywords_str = input("Введите ключевые слова через запятую: ")
    keywords = [kw.strip() for kw in keywords_str.split(',')]

    # Инициализация драйвера браузера
    browser = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

    # Перебор каждой строки в DataFrame
    for index, row in df.iterrows():
        institute_name = str(row[1])
        report_url = str(row[3])
        # Вызов функции для поиска и скачивания PDF-отчетов
        download_pdf_reports(browser, institute_name, report_url, keywords)

except Exception as e:
    handle_exceptions(e)
finally:
    if browser is not None:  # Проверяем, определена ли переменная browser перед её закрытием
        browser.close()
