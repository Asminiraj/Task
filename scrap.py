import openpyxl
from bs4 import BeautifulSoup
import requests

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title='Place'
sheet.append(['Country', 'Population', 'Year', 'Net Worth','Density'])
try:
    response = requests.get(
        'https://www.worldometers.info/world-population/population-by-country/'
    )

    soup = BeautifulSoup(response.text, "html.parser")

    table_body = soup.find('tbody')

    rows = table_body.find_all('tr')

    for row in rows:
        cols = row.find_all('td')

        country = cols[1].text.strip()
        population = cols[2].text.strip()
        year = cols[3].text.strip()
        net_worth = cols[4].text.strip()
        Density = cols[5].text.strip()
        sheet.append([
            country,
            population,
            year,
            net_worth,
            Density
        ])
        #print(country, population,year,net_worth)
except Exception as e:
    print(e)
excel.save("Place.xlsx")