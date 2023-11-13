import re
from collections import OrderedDict
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from common import column_to_letter, curly_brace, current_URL, get_options, get_driver, Clipboard
import time
from openpyxl import Workbook, worksheet
from openpyxl.chart import LineChart, Reference

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


def insert_newsheet(clip: Clipboard):
    new_sheetname = 'newSheet'

    # type: Workbook
    wb = clip.wb

    # type: worksheet
    ws = wb.create_sheet(new_sheetname)
    new_sheet_index = len(wb.sheetnames)-1
    wb.active = new_sheet_index
    # ws = self.wb.active  # type: worksheet

    # Iterate to ~60 rows
    # len 58
    for i in range(1, 60):
        # TODO Tuple error: self[key].value = value
        # AttributeError: 'tuple' object has no attribute 'value'
        # ws['A{index}:B{index}'.format(index=i)] = '=Income!A{index}:B{index}'.format(index=i)
        ws.cell(column=1, row=i, value='=Income!{letter}{index}'.format(index=i, letter=column_to_letter(1)))

    # Column tuple for reference, new index, and text ref in Excel.
    # TODO Industry specific ration CFO / Capex
    # https://www.investopedia.com/terms/c/capitalexpenditure.asp
    col_tuples = [{'ref': 2, 'idx': 2, 'text': '', 'table': 'Income'},
                  {'ref': 4, 'idx': 4, 'text': 'Gross', 'table': 'Income'},
                  {'ref': 9, 'idx': 6, 'text': 'Operating', 'table': 'Income'},
                  {'ref': 17, 'idx': 8, 'text': 'Net', 'table': 'Income'},
                  {'ref': 11, 'idx': 10, 'text': 'OCF', 'table': 'Cash'},
                  # Ignoring flag in Net change in PPE
                  {'ref': 12, 'idx': 12, 'text': 'Capex', 'table': 'Cash', 'ignore': True},
                  ]
    for t in col_tuples:
        for i in range(1, 60):
            ws.cell(column=t['idx'], row=i, value='={table}!{letter}{index}'
                    .format(index=i, letter=column_to_letter(t['ref']), table=t['table']))

    # TODO hard code to rev_col
    rev_col = 2
    ws.cell(column=rev_col+1, row=1, value='Revenue growth %')
    for i in range(1, 60):
        if (i-4) > 1:
            ws.cell(column=rev_col+1, row=i, value='=({letter}{index}-{letter}{index2})/{letter}{index2}'
                    .format(index=i, index2=i-4, letter=column_to_letter(rev_col)))
            ws['{letter}{index}'.format(index=i, letter=column_to_letter(rev_col+1))].number_format = '0.00%'

    # TODO Free cash burn
    fcf_col = 13    # Allocate M for free cash flow
    fcf_margin_col = fcf_col+1
    ocf_letter = 'J'    # OCF
    net_ppe_letter = 'L'    # Net PPE
    ws.cell(column=fcf_col, row=1, value='FCF')
    ws.cell(column=fcf_margin_col, row=1, value='FCF margin %')
    for i in range(2, 60):
        ws.cell(column=fcf_col, row=i, value='={ocf_letter}{index}+{net_ppe_letter}{index}'
                .format(ocf_letter=ocf_letter, net_ppe_letter=net_ppe_letter, index=i))
        ws.cell(column=fcf_col+1, row=i, value='={fcf_letter}{index}/B{index}'
                .format(fcf_letter=column_to_letter(fcf_col), index=i))
        ws['N{}'.format(i)].number_format = '0.00%'

    for t in col_tuples:
        if t['text'] != '':
            if 'ignore' in t:
                continue
            ws.cell(column=t['idx']+1, row=1, value='{} margin %'.format(t['text']))
            for i in range(2, 60):
                ws.cell(column=t['idx']+1, row=i,
                        value='={letter1}{index}/{letter2}{index}'
                        # TODO Fixed the reference to column B
                        .format(index=i, letter1=column_to_letter(t['idx']), letter2='B'))  # B - Revenue col
                ws['{letter}{index}'.format(index=i, letter=column_to_letter(t['idx']+1))].number_format = '0.00%'

    # Graph data
    chart = LineChart()
    letter = column_to_letter(rev_col+1)
    data = Reference(ws, range_string=f'{new_sheetname}!{letter}1:{letter}60')
    chart.add_data(data, titles_from_data=True)
    for t in col_tuples:
        if t['text'] != '':
            letter = column_to_letter(t['idx']+1)
            if 'ignore' in t:
                continue
            data = Reference(ws, range_string=f'{new_sheetname}!{letter}1:{letter}60')
            chart.add_data(data, titles_from_data=True)

    # Adding FCF margin column in chart
    letter = column_to_letter(fcf_margin_col)
    data = Reference(ws, range_string=f'{new_sheetname}!{letter}1:{letter}60')
    chart.add_data(data, titles_from_data=True)

    category = Reference(ws, range_string=f'{new_sheetname}!A2:A60')
    chart.set_categories(category)
    ws.add_chart(chart, 'D3')


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

            insert_newsheet(clip)
            clip.save()
            driver.quit()


if '__main__' == __name__:
    run_main()
