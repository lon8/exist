import requests
from bs4 import BeautifulSoup
import openpyxl
import xlsxwriter
# from selenium import webdriver


workbook = openpyxl.load_workbook('result.xlsx')

worksheet = workbook.active

wb = xlsxwriter.Workbook('test.xlsx')
ws = wb.add_worksheet()

value = worksheet.cell(row=11, column=2).value

# browser = webdriver.Chrome('chromedriver.exe')

# browser.get(f'https://exist.ru/Price/?pcode={value}')

with open('index.html', 'r', encoding='utf-8') as file:
    src = file.read()

soup = BeautifulSoup(src, 'lxml')

brand = soup.find('div', class_='art').text

print(brand)

offers = soup.find('div', class_='allOffers').find_all('div', class_='pricerow')

counter = 0

for offer in offers:
    avail_bool = True
    avail = offer.find('div', class_='avail').find('a', class_='gal')
    if avail is None:
        avail_bool = False
    delivery_date = offer.find('span', class_='statis').text
    price = offer.find('span', class_='price').text
    price = ''.join([x for x in price if x.isdigit()])
    print(price)
    ws.write(counter, 0, brand)
    ws.write(counter, 1, value ) # Артикул
    ws.write(counter, 2, avail_bool)
    ws.write(counter, 3, delivery_date)
    ws.write(counter, 4, price)
    counter += 1

workbook.close()
wb.close()