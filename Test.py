# Модуль os предоставляет функционал для работы с операционной системой, включая управление процессами, файловой системой и окружением.
import os

# Модуль time используется для работы со временем, включая получение текущего времени, ожидание определённого интервала и другие операции.
import time

# Класс Options из модуля selenium.webdriver.chrome.options используется для настройки параметров запуска браузера Chrome при использовании Selenium.
from selenium.webdriver.chrome.options import Options

# Модуль pandas предоставляет высокоуровневые структуры данных и функции для работы с данными, такие как DataFrame и Series.
import pandas as pd

# Этот импорт позволяет создавать экземпляры веб-драйвера, такие как chrome, firefox, ie и т.д., для управления браузерами через Selenium.
from selenium import webdriver

# Эти исключения используются для обработки ситуаций, когда элемент не найден или ссылка на элемент стала устаревшей в Selenium.
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException

# Класс By содержит константы, которые определяют способы поиска элементов на веб-странице в Selenium.
from selenium.webdriver.common.by import By

# Этот модуль предоставляет класс для удалённых веб-элементов в Selenium.
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement

# Этот импорт, относится к пользовательскому модулю, который обрабатывает пути к файлам или директориям.
import pathes

# Модуль datetime предоставляет классы для работы с датами и временем в Python.
from datetime import datetime

from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By


#Переменные для задани диапазона поиска на сайте
start_site = 1678707
end_site = start_site + 255
# end_site = 45977960

#Задание времени создания базы
now = datetime.now()
current_time = now.strftime("%d.%m %H_%M")
print("Current Time =", current_time)

#Описание столбцов для вывода в excel
TABLE_COLUMNS = ['Ссылка на компанию', 'Название компании', 'Почта 1', 'Почта 2', 'Почта 3', 'Почта 4', 'ИНН', 'КПП', 'Фактический адрес', 'Юридический адрес', 'Телефон 1', 'Телефон 2', 'Телефон 3', 'Телефон 4', 'Генеральный директор', 'ОКВЭД2', 'Дата регистрации', 'Налоговые отчисления', 'Выручка за 2023 год', 'Сайт' ]
TABLE = {column: [] for column in TABLE_COLUMNS}

#Подключение к браузеру
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
# .\chrome.exe --remote-debugging-port=9222 --user-data-dir="~/ChromeProfile"
browser = webdriver.Chrome(options=chrome_options)

#Функия получения текста
def get_element_text(driver: WebDriver, path: str) -> str:
    try:
        return driver.find_element(By.XPATH, path).text
    except NoSuchElementException:
        return ''

#Функция
def move_to_element(driver: WebDriver, element: WebElement | WebDriver) -> None:
    try:
        webdriver.ActionChains(driver).move_to_element(element).perform()
    except StaleElementReferenceException:
        pass

#Функция нажатия на элемент
def element_click(driver: WebDriver | WebElement, path: str) -> bool:
    try:
        driver.find_element(By.XPATH, path).click()
        return True
    except:
        return False

def delete_cache():
    browser.get('chrome://settings/clearBrowserData')
    time.sleep(1)
    actions = ActionChains(browser)
    actions.send_keys(Keys.TAB + Keys.ENTER) # выбираем удаление кэша
    actions.perform()
    time.sleep(1)

def Login():
    browser.get('https://i.moscow/connect/sudir')
    time.sleep(5)
    element_click(browser, pathes.Aut1)
    time.sleep(5)
    element_click(browser, pathes.Aut2)
    time.sleep(5)
    element_click(browser, pathes.Aut3)
    time.sleep(5)

#основная программа
def main():
    # Инициализация счетчика
    counter = 0

    # Предполагаемый текст для проверки
    expected_text = 'nginx'

    for x in range(start_site, end_site):

        print(counter)
        # Выполнение команды каждые 50 итераций
        if counter % 50 == 0:
            delete_cache()
            Login()
            print('Start')

        # Увеличение счетчика
        counter += 1
        time.sleep(0.1)

        url = 'https://i.moscow/company/' + str(x)
        print(url)
        # os.startfile('.\chrome.exe --remote-debugging-port=9222 --user-data-dir="~/ChromeProfile"')
        # time.sleep(500)
        # Подключение к браузеру
        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        # .\chrome.exe --remote-debugging-port=9222 --user-data-dir="~/ChromeProfile"
        browser = webdriver.Chrome(options=chrome_options)
        browser.get(url)
        time.sleep(0.1)


        checking_text = get_element_text(browser, pathes.Check)
        print(checking_text)
        if checking_text != expected_text:

            name_company = get_element_text(browser, pathes.name_company)
            inn_company = get_element_text(browser, pathes.inn_company)
            kpp_company = get_element_text(browser, pathes.kpp_company)
            email1 = get_element_text(browser, pathes.email1)
            email2 = get_element_text(browser, pathes.email2)
            email3 = get_element_text(browser, pathes.email3)
            email4 = get_element_text(browser, pathes.email4)
            f_addres = get_element_text(browser, pathes.f_addres)
            y_addres = get_element_text(browser, pathes.y_addres)
            tel1 = get_element_text(browser, pathes.tel1)
            tel2 = get_element_text(browser, pathes.tel2)
            tel3 = get_element_text(browser, pathes.tel3)
            tel4 = get_element_text(browser, pathes.tel4)
            gendir = get_element_text(browser, pathes.gendir)
            okved2 = get_element_text(browser, pathes.okved2)
            date_reg = get_element_text(browser, pathes.date_reg)
            nalog_ot = get_element_text(browser, pathes.nalog_ot)
            viruchka = get_element_text(browser, pathes.viruchka)
            site = get_element_text(browser, pathes.site)

            print('ИНН: ' + inn_company)
            print('КПП: ' + kpp_company)
            print('Email: ' + email1)

            TABLE['Ссылка на компанию'].append(url)
            TABLE['Название компании'].append(name_company)
            TABLE['Почта 1'].append(email1)
            TABLE['Почта 2'].append(email2)
            TABLE['Почта 3'].append(email3)
            TABLE['Почта 4'].append(email4)
            TABLE['ИНН'].append(inn_company)
            TABLE['КПП'].append(kpp_company)
            TABLE['Фактический адрес'].append(f_addres)
            TABLE['Юридический адрес'].append(y_addres)
            TABLE['Телефон 1'].append(tel1)
            TABLE['Телефон 2'].append(tel2)
            TABLE['Телефон 3'].append(tel3)
            TABLE['Телефон 4'].append(tel4)
            TABLE['Генеральный директор'].append(gendir)
            TABLE['ОКВЭД2'].append(okved2)
            TABLE['Дата регистрации'].append(date_reg)
            TABLE['Налоговые отчисления'].append(nalog_ot)
            TABLE['Выручка за 2023 год'].append(viruchka)
            TABLE['Сайт'].append(site)

            time.sleep(0.1)
            browser.quit()
            pd.DataFrame(TABLE).to_excel(f"{'База компаний Инновационного Кластера Москвы '+current_time}.xlsx")

        else:
            print('Сработала защита nginx')
            delete_cache()
            Login()


if __name__ == '__main__':
    main()