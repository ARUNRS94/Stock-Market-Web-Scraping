import requests
import bs4
from bs4 import BeautifulSoup
import xlwings as xw
from playsound import playsound

def getvalue():
    r=requests.get('https://finance.yahoo.com/quote/TATAMOTORS.NS?p=TATAMOTORS.NS')
    s=bs4.BeautifulSoup(r.text,"lxml")
    value = s.find_all("div", {'class': 'My(6px) Pos(r) smartphone_Mt(6px)'})[0].find('span').text
    return value

while True:
    livevalue = float(getvalue())
    wb = xw.Book('Stock_Value_alert.xlsx')
    sheet = wb.sheets['Sheet1']
    sheet.range('B2').value = livevalue
    alertvalue = sheet.range('B3').value
    if livevalue < alertvalue:
        playsound('F:/ArunRS/Internship/Stock Market Web Scraping/ding-sound.mp3')
        sheet.range('B4').value= "Price Reached"
    else:
        difference= livevalue-alertvalue
        sheet.range('B4').value=difference

    
