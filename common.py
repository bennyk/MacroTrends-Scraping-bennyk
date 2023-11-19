from openpyxl import Workbook, worksheet
from openpyxl.utils.dataframe import dataframe_to_rows
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import re
import pandas as pd
from openpyxl import Workbook, worksheet
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Font, Alignment

from itertools import tee
from typing import List, Dict


def pairwise(iterable):
    "s -> (s0,s1), (s1,s2), (s2, s3), ..."
    a, b = tee(iterable)
    next(b, None)
    return zip(a, b)


def isnumeric(a: str):
    return re.match(r'[.0-9]+$', a)


def get_driver(options):
    driver = webdriver.Firefox(executable_path=r'C:\Users\benny\PyCharmProjects\MacroTrends-Scraping\geckodriver.exe',
                               options=options)
    return driver


def get_options():
    options = Options()
    options.binary_location = r'C:\Users\benny\AppData\Local\Mozilla Firefox\firefox.exe'
    return options


def current_URL(url_path):
    options = get_options()
    driver = get_driver(options)
    driver.get(url_path)
    box = driver.find_element(By.CSS_SELECTOR, ".js-typeahead")
    ticker = 'intc'
    box.send_keys(ticker)
    time.sleep(1)
    box.send_keys(Keys.DOWN, Keys.RETURN)
    time.sleep(1)
    current_url = driver.current_url
    time.sleep(4)
    driver.quit()
    return current_url


def parse_grid(driver, fin_url):
    driver.get(fin_url)
    driver.set_window_size(2000, 2000)
    time.sleep(4)
    arrow = driver.find_element(By.CSS_SELECTOR, ".jqx-icon-arrow-right")
    webdriver.ActionChains(driver).click_and_hold(arrow).perform()
    # webdriver.ActionChains(driver).click_and_hold(arrow).move_by_offset(-1500, 0).release().perform()
    time.sleep(4)
    column_grid = driver.find_element(By.CSS_SELECTOR, "#columntablejqxgrid").text

    return parse_content(driver, column_grid)


def parse_content(driver, col_grid):
    # prune it for $ and comma
    content_grid = driver.find_element(By.CSS_SELECTOR, "#contenttablejqxgrid").text
    content_grid = content_grid.replace("$", "").replace(",", "")
    content_grid = content_grid.splitlines()
    data_dict = {'Years': col_grid.split('\n')[1:]}
    last_key = None
    for s in content_grid:
        if not (isnumeric(s) or s == '-' or s.startswith('-')):
            last_key = s
        else:
            if last_key is None:
                continue  # Skip data if no key is available

            try:
                ticker = float(s)
                data_dict.setdefault(last_key, []).append(ticker)
            except ValueError:
                data_dict.setdefault(last_key, []).append(s)
    return data_dict


def curly_brace(my_str, open_list=None, close_list=None):
    # Function to check parentheses

    # TODO Could not solve
    # for a in re.finditer(r'\[([^]]+)\]\[([^]]*)\]', g):

    if close_list is None:
        close_list = ["}"]
    if open_list is None:
        open_list = ["{"]

    def _check(str):
        # https://www.geeksforgeeks.org/check-for-balanced-parentheses-in-python/
        stack = []
        string = ''
        level = 0
        for c in str:
            if c in open_list:
                stack.append(c)
                level += 1
                if level == 1:
                    string = ''
            elif c in close_list:
                pos = close_list.index(c)
                if ((len(stack) > 0) and
                        (open_list[pos] == stack[len(stack) - 1])):
                    stack.pop()
                    level -= 1
                    if level == 0:
                        # print(string)
                        yield string
                else:
                    return "Unbalanced"
            else:
                string += c
        if len(stack) == 0:
            return "Balanced"
        else:
            return "Unbalanced"

    return _check(my_str)


class MyVar:
    def __init__(self):
        self.var = 1

    def gen_var(self):
        self.var += 1
        return self.var


class Clipboard:
    def __init__(self):
        self.wb = Workbook()
        self.excel_file = 'revenue_and_profit.xlsx'

        # removing initial sheet
        ws = self.wb.active
        self.wb.remove(ws)

    def write_excel(self, title, df):
        ws = self.wb.create_sheet(title)

        # Add the column headers
        for row in dataframe_to_rows(df, index=False, header=True):
            ws.append(row)

    def write_dict_to_excel(self, title, data_dict):
        df = pd.DataFrame(data_dict)
        df = df.replace('-', 0)
        df = df.sort_values(by='Years', ascending=True)
        ws = self.wb.create_sheet(title)

        # Add the column headers
        for row in dataframe_to_rows(df, index=False, header=True):
            ws.append(row)

    def save(self):
        # Save the Excel file
        self.wb.save(self.excel_file)

    def insert_newsheet(self, new_sheetname: str, col_tuples: List[Dict]):
        ws = self.wb.create_sheet(new_sheetname)
        new_sheet_index = len(self.wb.sheetnames)-1
        self.wb.active = new_sheet_index

        # Iterate to ~60 rows
        # len 58
        for i in range(1, 60):
            ws.cell(column=1, row=i, value='=Income!{letter}{index}'.format(index=i, letter=column_to_letter(1)))

        for t in col_tuples:
            for i in range(1, 60):
                ws.cell(column=t['idx'], row=i, value='={table}!{letter}{index}'
                        .format(index=i, letter=column_to_letter(t['ref']), table=t['table']))
        return ws

    def insert_chart(self, ws, start, end, new_sheetname):

        # TODO insert start result col and end result col
        for j in range(1, end):
            ws.column_dimensions[column_to_letter(j)].width = 10
            ws.cell(column=j, row=1).alignment = Alignment(wrapText=True)

        # Graph data
        chart = LineChart()
        for i in range(start, end):
            letter = column_to_letter(i)
            data = Reference(ws, range_string=f'{new_sheetname}!{letter}1:{letter}60')
            chart.add_data(data, titles_from_data=True)

        category = Reference(ws, range_string=f'{new_sheetname}!A2:A60')
        chart.set_categories(category)
        ws.add_chart(chart, 'D3')

    def insert_debt(self):
        var = MyVar()
        cash_col = var.gen_var()
        lt_assets_col = var.gen_var()
        total_current_liab_col = var.gen_var()
        lt_liabilities_col = var.gen_var()
        equity_col = var.gen_var()
        lt_liabilities_to_assets_col = var.gen_var()
        net_debt_to_equity_col = var.gen_var()
        cash_to_short_term_col = var.gen_var()

        col_tuples = [
            # Cash
            {'ref': 2, 'idx': cash_col, 'text': 'LT asset', 'table': 'Balance'},
            # LT assets
            {'ref': 12, 'idx': lt_assets_col, 'text': 'LT asset', 'table': 'Balance'},
            # ST debt
            {'ref': 14, 'idx': total_current_liab_col, 'text': 'LT debt', 'table': 'Balance'},
            # LT liabilities
            {'ref': 17, 'idx': lt_liabilities_col, 'text': 'LT liabilities', 'table': 'Balance'},
            # Shareholder equity
            {'ref': 23, 'idx': equity_col, 'text': 'Equity', 'table': 'Balance'},
        ]
        new_sheetname = 'Debt'
        ws = self.insert_newsheet(new_sheetname, col_tuples)

        ws.cell(column=lt_liabilities_to_assets_col, row=1, value='Long term liabilities to assets')
        ws.cell(column=net_debt_to_equity_col, row=1, value='Net debt to equity')
        ws.cell(column=cash_to_short_term_col, row=1, value='Cash to short-term borrowing')
        for i in range(2, 60):
            ws.cell(column=lt_liabilities_to_assets_col, row=i, value='=${lt_liabilities}{index}/${lt_assets}{index}'
                    .format(index=i,
                            lt_liabilities=column_to_letter(lt_liabilities_col),
                            lt_assets=column_to_letter(lt_assets_col)))
            ws['{}{}'.format(column_to_letter(lt_liabilities_to_assets_col), i)].number_format = '0%'

            ws.cell(column=net_debt_to_equity_col, row=i,
                    value='=(${total_lt_liab}{index}+${total_current_liab}{index}-${cash}{index})/${equity}{index}'
                    .format(index=i,
                            cash=column_to_letter(cash_col),
                            total_current_liab=column_to_letter(total_current_liab_col),
                            total_lt_liab=column_to_letter(lt_liabilities_col),
                            equity=column_to_letter(equity_col), ))
            ws['{}{}'.format(column_to_letter(net_debt_to_equity_col), i)].number_format = '0%'

            ws.cell(column=cash_to_short_term_col, row=i,
                    value='=${cash}{index}/${total_current_liab}{index}'
                    .format(index=i,
                            cash=column_to_letter(cash_col),
                            total_current_liab=column_to_letter(total_current_liab_col), ))
            ws['{}{}'.format(column_to_letter(cash_to_short_term_col), i)].number_format = '0.00'

        self.insert_chart(ws, lt_liabilities_to_assets_col, cash_to_short_term_col + 1,
                          new_sheetname)


def old_run_main():
    main_url_path = 'https://www.macrotrends.net/'
    current_url = current_URL(main_url_path)

    if "stocks" in current_url:
        # Check if the data in the ticker is available
        url_parts = current_url.split("/", 10)
        url_path = main_url_path+"stocks/charts/"+url_parts[5]+"/"+url_parts[6]+"/"
        driver = get_driver(get_options())
        # financial-statements
        fin_url_path = url_path+"financial-statements"
        driver.get(fin_url_path)
        if driver.find_elements(
                By.CSS_SELECTOR,
                "div.jqx-grid-column-header:nth-child(1) > div:nth-child(1) > div:nth-child(1) > span:nth-child(1)"):

            clip = Clipboard()

            income_url = url_path+"income-statement"
            data_dict = parse_grid(driver, income_url)
            clip.write_excel('Income', data_dict)

            balance_url = url_path+"balance-sheet"
            data_dict = parse_grid(driver, balance_url)
            clip.write_excel('Balance', data_dict)

            cash_url = url_path+"cash-flow-statement"
            data_dict = parse_grid(driver, cash_url)
            clip.write_excel('Cash', data_dict)

            clip.save()
            driver.quit()


def column_to_letter(column_int):
    # https://stackoverflow.com/questions/23861680/convert-spreadsheet-number-to-column-letter
    start_index = 1   # It can start either at 0 or at 1
    letter = ''
    while column_int > 25 + start_index:
        letter += chr(65 + int((column_int-start_index)/26) - 1)
        column_int = column_int - (int((column_int-start_index)/26))*26
    letter += chr(65 - start_index + (int(column_int)))
    return letter


def letter_to_column(name):
    # https://cwestblog.com/2013/09/13/python-convert-excel-column-name-to-number/
    pow = 1
    column_int = 0
    for letter in name[::-1]:
            column_int += (int(letter, 36) -9) * pow
            pow *= 26
    return column_int
