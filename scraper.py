import re
from collections import OrderedDict
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from common import column_to_letter, curly_brace, current_URL, get_options, get_driver, Clipboard
import time
from openpyxl import Workbook, worksheet
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Font, Alignment

# Inspired by https://github.com/capuccino26/MacroTrends-Scraping
# DataFrame https://www.geeksforgeeks.org/python-pandas-dataframe/
# https://regex101.com/
# https://www.debuggex.com/cheatsheet/regex/python
# https://pythex.org/


def extract(orig_data):
    # Pull out statements
    m = re.search(r'\[(.*)\]', orig_data)
    if m.group():
        # print(m.group(1))

        # Extract out QoQ/YoY statement
        # data = {
        #   "Years": ['2023-12-31', '2023-09-30', ],
        #   "Revenue": [50, 40, 45]
        # }

        # Init
        od = OrderedDict()
        first = True
        arr = None

        # Number of years
        # print(len(od['Years']))
        years = od['Years'] = []

        for s in curly_brace(m.group(1)):
            period = []
            for p in re.finditer(r'"((?:[^"\\]|\\.)*)"|([.\-\d]+)', s):
                period.append(p)
                if len(period) == 2:
                    # Match the pattern [field_name][<a href='...>][date1][value1][date2][value2]...
                    # print(period[0].group(), period[1].group())
                    field = period[0]
                    para = period[1]
                    if field.group(1) == 'field_name':
                        # Begin of statement
                        tag = re.match(r'<[^>]+>(.+)<[^>]+>', para.group(1))
                        assert tag is not None

                        # Collect the data based on array of dict based on period
                        arr = od[tag.group(1)] = []

                        if first and len(years) != 0:
                            first = False

                    elif re.match(r'\d{4}-\d{2}-\d{2}', field.group(1)):
                        if first:
                            # First timer
                            years.append(field.group(1))

                        # Date
                        val = 0.    # Default to zero
                        if para.group() != '':
                            # Extract out the value
                            if para.group(1) is not None and para.group(1) != '':
                                val = float(para.group(1))
                            elif para.group(2) is not None:
                                # .group(2) match unquoted bare value occasionally.
                                val = float(para.group(2))
                            assert val is not None
                        arr.append(val)

                    elif field.group(1) == 'popup_icon':
                        pass
                    else:
                        assert False
                    period.clear()
        # print(od)

        # Convert od to DataFrame
        last = None
        for f, v in od.items():
            # print(f)
            if last is None:
                last = len(v)
                # print("First len", last)
            else:
                if last != len(v):
                    print("{}: diff length {}".format(f, len(v)))
                # else:
                #     print("Matched", f)

        df = pd.DataFrame(od)
        df = df.sort_values('Years')
        # print(df)
        return df


def get_page(driver: webdriver, fin_url):
    driver.get(fin_url)
    driver.set_window_size(2000, 2000)
    time.sleep(4)

    data = driver.page_source
    for line in data.split('\n'):
        if re.search('var\s+originalData', line):
            df = extract(line)
            return df
    return None


def run_main():
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

            income_url = url_path+"income-statement?freq=Q"
            df = get_page(driver, income_url)
            clip.write_excel('Income', df)

            balance_url = url_path+"balance-sheet?freq=Q"
            df = get_page(driver, balance_url)
            clip.write_excel('Balance', df)

            cash_url = url_path+"cash-flow-statement?freq=Q"
            df = get_page(driver, cash_url)
            clip.write_excel('Cash', df)

            clip.insert_income()
            clip.insert_debt()
            clip.insert_returns()
            clip.insert_eps()
            clip.insert_shares_out()

            clip.save()
            driver.quit()


if '__main__' == __name__:
    run_main()
