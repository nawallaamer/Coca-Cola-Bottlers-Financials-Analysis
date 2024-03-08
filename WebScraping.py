import requests
from bs4 import BeautifulSoup
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains

print("Welcome to the Analysis by Nawall.")
# ticker = input('Enter the Required Company Name: ')
# ticker = str(ticker)
tickers = ['MX/XMEX/AC','TR/XIST/AEFES.E','TR/XIST/CCOLA.E','KO','KOF','CCHGY','BUD','PEP','CCEP','AKOA']
#Functions for conversion into USD
lira_er=[16.5754, 8.8922, 7.0194, 5.6828, 4.8456]
mexican_er=[20.1077, 20.2853, 21.48, 19.25, 19.22]
euro_er=[0.951, 0.8458, 0.877, 0.8931, 0.8475]
pound_er=[0.8115, 0.7271, 0.7798, 0.7835, 0.7501]
clp_er=[874.4487, 760.3627, 792.0704, 703.6885, 642.1589]
def Lira(val, er):
    if val.endswith('%'):
        return val
    else:
        return round(float(val.replace(',',''))/er,2)
def mexican(val1, er1):
    if val1.endswith('%'):
        return val1
    else:
        return round(float(val1.replace(',',''))/er1,2)
def euro(val1, er2):
    if val1.endswith('%'):
        return val1
    else:
        return round(float(val1.replace(',',''))/er2,2)
def pound(val1, er3):
    if val1.endswith('%'):
        return val1
    else:
        return round(float(val1.replace(',',''))/er3,2)
def clp(val1, er4):
    if val1.endswith('%'):
        return val1
    else:
        return round(float(val1.replace(',',''))/er4,2)

for ticker in tickers:
    webdriver_path = r'C:\COKE PROJ\cd\chromedriver.exe'
    service = Service(webdriver_path)

    # Configure Selenium options
    options = Options()
    # Avoid opening a visible browser window
    options.add_argument("--headless") 
    driver = webdriver.Chrome(service=service, options=options)

    header = {
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64 x64; rv: 87.0) Gecko/20100101 FireFox/87.0',
        'Accept': 'text/html,application/xhtml+xml, application/xml; q=0.9, image/webp, /; q=0.8',
        'Accept-Language': 'en-US,en; q=0.5',
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1',
        'Cache-Control': 'max-age=0'
    }

    url = {}
    url['Income Statement'] = f"https://www.wsj.com/market-data/quotes/{ticker}/financials/annual/income-statement"
    url['Balance Sheet'] = f"https://www.wsj.com/market-data/quotes/{ticker}/financials/annual/balance-sheet"
    url['Cash Flow'] = f"https://www.wsj.com/market-data/quotes/{ticker}/financials/annual/cash-flow"

    if ticker == 'MX/XMEX/AC':
        ticker1 = 'AC.MX'
    elif ticker == 'TR/XIST/AEFES.E':
        ticker1 = 'AEFES.E'
    elif ticker == 'TR/XIST/CCOLA.E':
        ticker1 = 'CCOLA.E'
    else:
        ticker1 = ticker

    xl = pd.ExcelWriter(f'Financial Statements ({ticker1}).xlsx', engine='xlsxwriter')
    for key, url in url.items():
        driver.get(url)

        html = driver.page_source

        soup = BeautifulSoup(html, 'html.parser')

        tables = soup.find_all('table', {'class': 'cr_dataTable'}, recursive=True)
        rows = []
        for table in tables:
            for row in table.find_all('tr'):  # Find all rows within each table
                row_data = [cell.get_text(strip=True) for cell in row.find_all(['th', 'td'])]
                if not row_data or all(cell == '' for cell in row_data) or row_data[1:6] == ['-','-','-','-','-']:
                    continue
                else:
                    if row_data[0] in ['All values MXN Millions.', 'All values TRY Millions.', 'All values USD Millions.','All values GBP Thousands.']:
                        continue
                    else:
                        row_data = [f'-{cell[1:-1]}' if cell.startswith('(') and cell.endswith(')') else cell for cell in row_data]
                        row_data = ['0' if cell == '-' else cell for cell in row_data]
                        rows.append(row_data[0:6])

        columns = rows[0]
        rows = rows[1:]

        df = pd.DataFrame(rows, columns=columns)
        df.columns.values[0]='Fiscal year is January-December. All values USD Millions.'
        print(df)

        #Conversion into USD
        if ticker1=='AEFES.E' or ticker1== 'CCOLA.E':
            for r in range(len(lira_er)):
                datecol=f'{2022-r}'
                er=lira_er[r]
                df[datecol]=df[datecol].apply(Lira,er=er)
        elif ticker1== 'AC.MX' or ticker1== 'KOF':
            for r in range(len(mexican_er)):
                datecol1=f'{2022-r}'
                er1=mexican_er[r]
                df[datecol1]=df[datecol1].apply(mexican, er1=er1)
        elif ticker1== 'AKOA':
            for r in range(len(clp_er)):
                datecol1=f'{2022-r}'
                er4=clp_er[r]
                df[datecol1]=df[datecol1].apply(clp, er4=er4)
        elif ticker1== 'BUD':
            for r in range(len(euro_er)):
                datecol1=f'{2022-r}'
                er2=euro_er[r]
                df[datecol1]=df[datecol1].apply(euro, er2=er2)
        elif ticker1== 'CCHGY':
            for r in range(len(pound_er)):
                datecol1=f'{2022-r}'
                er3=pound_er[r]
                df[datecol1]=df[datecol1].apply(pound, er3=er3)

        df.to_excel(xl, sheet_name=key, index=False)
        worksheet = xl.sheets[key]

        for r, column in enumerate(df.columns):
            column_width = max(df[column].astype(str).map(len).max(), len(column))
            worksheet.set_column(r, r, column_width)

    xl.close()
    driver.quit()
