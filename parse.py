import time

from selenium import webdriver
from selenium.webdriver.common.by import By

import sqlite3

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill

# назви файлів
EXCEL_FILENAME = 'today_rate.xlsx'
DATABASE_FILENAME = 'rates.db'

# 
def get_exchange_rate():
    driver = webdriver.Chrome()
    driver.get("https://www.google.com/finance/quote/USD-UAH")

    element = driver.find_element(By.CLASS_NAME,'fxKbKc')
    course = element.text
    
    driver.close()

    print(f'Поточний курс USD/UAH: {course}')
    return course


def database_start():
    db = sqlite3.connect(DATABASE_FILENAME)
    cursor = db.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS exchange_rates (
            id INTEGER PRIMARY KEY,
            datetime TEXT NOT NULL,
            rate REAL NOT NULL
        )
    ''')
    db.commit()
    db.close()


def database_add_rate(datetime_sql, rate):
    db = sqlite3.connect(DATABASE_FILENAME)
    cursor = db.cursor()

    cursor.execute('INSERT INTO exchange_rates (datetime, rate) VALUES (?, ?)', (datetime_sql, rate))

    db.commit()
    db.close()


def database_get_data_for_excel():
    db = sqlite3.connect(DATABASE_FILENAME)
    cursor = db.cursor()
    
    cursor.execute('''
        SELECT strftime('%d.%m.%Y %H:%M:%S', datetime), rate
        FROM exchange_rates
        WHERE date('now') = date(datetime)
    ''')

    today_rates = cursor.fetchall()

    db.commit()
    db.close()

    return today_rates


def excel_start():
    try:
        wb = openpyxl.load_workbook(EXCEL_FILENAME)
    except:
        wb = Workbook()
        sheet = wb.active
        
        sheet.column_dimensions['A'].width = 20
        sheet.column_dimensions['B'].width = 15
        
        fill = PatternFill(fill_type='solid', fgColor='00FFFF00')
        sheet['A1'].fill = fill
        sheet['B1'].fill = fill
        
        wb.save(EXCEL_FILENAME)
        wb.close()


def update_rates_in_excel(filename, today_rates):
    wb = openpyxl.load_workbook(filename)
    sheet = wb.active

    for row in sheet.iter_rows(min_row=1, max_col = 2, max_row = sheet.max_row):
        for cell in row:
            cell.value = None

    wb.save(filename)
    wb.close()

    wb = openpyxl.load_workbook(filename)
    sheet = wb.active

    sheet['A1'], sheet['B1'] = 'datetime', 'exchange_rate'
    for rate in today_rates:
        sheet.append(rate)

    wb.save(filename)
    wb.close()


def get_data_for_sql(current_time):
    current_time_sql = time.strftime('%Y-%m-%d %H:%M:%S', current_time)
    rate = get_exchange_rate()
    return current_time_sql, rate



database_start()
excel_start()

while True:
    time.sleep(1)
    current_time = time.localtime()    
    if current_time.tm_min != 0 and current_time.tm_sec % 20 == 0:
        
        current_time_sql, rate = get_data_for_sql(current_time)
        database_add_rate(current_time_sql, rate)

        today_rates = database_get_data_for_excel()
        update_rates_in_excel(EXCEL_FILENAME, today_rates)
