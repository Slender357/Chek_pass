from seleniumwire import webdriver
import gspread
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
import chromedriver_autoinstaller
import time
from datetime import datetime,timedelta
from config import CONFIG
from partools import answer_captcha, get_sheet_values, quickstart_sheet, safe_to_shet, chek_last_date, update_last_date
from bs4 import BeautifulSoup
import requests


def chek_pass(row_car: dict):
    all_pass = get_values(row_car['ГРЗ'])
    return all_pass


def open_chrome():
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(options=options,
                              executable_path=chromedriver_autoinstaller.install())
    driver.get(CONFIG['url'])
    headers = None
    for request in driver.requests:
        if 'https://reestr.ovga.mos.ru/generate-captcha/gcb_captcha' in request.url:
            headers = request.headers
    return driver, headers


def select_on_page(driver, car, headers, seria: str):
    while True:
        time.sleep(1)
        src = driver.find_element(By.XPATH, '/html/body/div[1]/form/div[2]/div/div[1]/img').get_attribute('src')
        p = requests.get(src, headers=headers)
        with open("img.png", "wb") as out:
            out.write(p.content)
        result = answer_captcha('img.png')
        select_object = Select(driver.find_element(By.XPATH, CONFIG['seria']))
        select_object.select_by_value(seria)
        driver.find_element(By.XPATH, CONFIG['car']).clear()
        driver.find_element(By.XPATH, CONFIG['car']).send_keys(car)
        driver.find_element(By.XPATH, CONFIG['captcha']).send_keys(result['captchaSolve'])
        time.sleep(1)
        driver.find_element(By.XPATH, CONFIG['buttom_search']).click()
        time.sleep(1)
        sp = BeautifulSoup(driver.page_source, 'html.parser').find('div', class_="alert alert-danger")
        if sp is not None:
            continue
        break



def get_values(car, driver, headers):
    select_on_page(driver, car, headers, 'БА')
    pars_tab = parsing_tab(driver, car)
    if test_pars_date(pars_tab):
        select_on_page(driver,car, headers, 'ББ')
        pars_tab = parsing_tab(driver, car, pars_tab)
    return pars_tab


def parsing_tab(driver, car, pars_tab=None):
    if pars_tab is None:
        pars_tab = {
            'car': car,
            'day': None,
            'night': None
        }
    sp = BeautifulSoup(driver.page_source, 'html.parser').select('table', class_="table table-hover")
    tr_info = sp[1].find_all('tr', class_='info')
    for i in tr_info:
        if pars_tab['day'] is None and i.find_all('td')[-1].text == 'Дневной':
            fdate = i.find_all('td')[3].text
            pars_tab['day'] = {
                'tdate': i.find_all('td')[2].text,
                'fdate': fdate,
                'type': i.find_all('td')[4].text,
            }
        if pars_tab['night'] is None and i.find_all('td')[-1].text == 'Ночной':
            fdate = i.find_all('td')[3].text
            pars_tab['night'] = {
                'tdate': i.find_all('td')[2].text,
                'fdate': fdate,
                'type': i.find_all('td')[4].text,
            }
    return pars_tab


def test_pars_date(pars_tab: dict) -> bool:
    """
    Тест словаря на отсутствие пропусков
    :param pars_tab:
    :return:
    """
    if pars_tab['day'] is None:
        return True
    elif pars_tab['night'] is None:
        return True
    else:
        return False


def chek_pass_bot():
    worksheet = gspread.service_account(filename=CONFIG['credentials_file']).open_by_key(CONFIG['spreadsheet_id']) \
        .worksheet(CONFIG['sheet'])
    spreadsheet_id, service = quickstart_sheet()
    if chek_last_date(spreadsheet_id, service):
        driver, headers = open_chrome()
        sheet_values = get_sheet_values(spreadsheet_id, service, CONFIG['sheet'])
        for i, values in enumerate(sheet_values):
            try:
                car = values['ГРЗ']
                pars_tab = get_values(car, driver, headers)
                safe_to_shet(car, worksheet, service, spreadsheet_id, i + 3, pars_tab)
            except BaseException:
                time.sleep(10)
                car = values['ГРЗ']
                pars_tab = get_values(car, driver, headers)
                safe_to_shet(car, worksheet, service, spreadsheet_id, i + 3, pars_tab)

        update_last_date(spreadsheet_id, service)


if __name__ == '__main__':
    chek_pass_bot()
