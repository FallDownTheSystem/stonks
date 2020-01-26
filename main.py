import datetime
import requests
from bs4 import BeautifulSoup
from dateutil.relativedelta import relativedelta
from openpyxl import Workbook
import sys
import shutil

try:
    market = sys.argv[1]
    underlying = sys.argv[2]
except Exception:
    pass

print(f'Type: {market}  Underlying: {underlying}')

months = {
    1: 'January',
    2: 'February',
    3: 'March',
    4: 'April',
    5: 'May',
    6: 'June',
    7: 'July',
    8: 'August',
    9: 'September',
    10: 'October',
    11: 'November',
    12: 'December'
}

# market = 'fund'
# underlying = 'SPY'
today = datetime.datetime.today()
filename: str = market + "_" + underlying + "_" + str(today)
filename = filename.replace(" ", "_").replace(".", "-").replace(":", "-") + ".xlsx"
print(f'Creating workbook')
stonks = Workbook()

for i in range(5):
    delta = relativedelta(months=i)
    date = today + delta
    month_name = months[date.month]
    year = date.year
    print(f'Creating worksheet {year}-{date.month}')
    if i == 0:
        ws = stonks.active
        ws.title = f'{year}-{date.month}'
    else:
        ws = stonks.create_sheet(title=f'{year}-{date.month}')

    URL = f'https://www.marketwatch.com/investing/{market}/{underlying}/OptionsMonth?month={month_name}&year={year}&countrycode=US'
    page = requests.get(URL)
    soup = BeautifulSoup(page.content, 'html.parser')
    rows = soup.find_all('tr')
    print("Parsing data")
    for row_i, row in enumerate(rows):
        cells = row.find_all('td')
        for col_i, cell in enumerate(cells):
            classes = cell.get("class")
            val = cell.text.strip()
            if classes is not None and not ("acenter" in classes or "aleft" in classes):
                val = val.replace(',', '')
                try:
                    val = float(val)
                except Exception:
                    pass

            ws.cell(column=col_i + 1, row=row_i + 1, value=val)

print(f"Saving to file {filename}")
stonks.save(filename=filename)
shutil.copy(filename, market + "_" + underlying + "_latest.xlsx")
