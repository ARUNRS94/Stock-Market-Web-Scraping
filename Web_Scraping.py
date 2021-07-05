import requests
import bs4
from bs4 import BeautifulSoup
import xlwings as xw
from playsound import playsound
from datetime import datetime

def getvalue(symbol):
 try:
    url= f'https://finance.yahoo.com/quote/{symbol}'
    r=requests.get(url, headers={'User-Agent': 'Custom'})
    print (symbol)
    s=bs4.BeautifulSoup(r.text,"lxml")
    value = s.find_all("div", {'class': 'My(6px) Pos(r) smartphone_Mt(6px)'})[0].find('span').text
    float_value = float(value.replace(',',''))
    return float_value
 except:
     return "Error"

def findlastrow(row):
    rowno = f'A{row}'
    emptyrow = sheet.range(rowno).value
    return emptyrow    

def priceupdate(cell):
    sym=f'A{cell}'
    liv=f'B{cell}'
    ale=f'C{cell}'
    dif=f'D{cell}'
    symbol=sheet.range(sym).value
    livevalue =getvalue(symbol)
    sheet.range(liv).value = livevalue
    alertvalue = sheet.range(ale).value
    difference= livevalue-alertvalue
    sheet.range(dif).value=difference   

wb = xw.Book('Stock_Value_alert.xlsx')
sheet = wb.sheets['Sheet1']    
row=0
rowvalue=""

while True:
    if rowvalue == None:
        last_row = row
        print(last_row)
        break
    else:
        row=row+1
        rowvalue=findlastrow(row)
        
cell = 2
finaltime= datetime.now()

while True:
    if last_row == cell:
        responce=finaltime -datetime.now()
        sheet.range('G1').value= responce.total_seconds()
        cell = 2
        finaltime= datetime.now()
    else:
        priceupdate(cell)
        cell=cell+1       
                  
'''while True:
    
    if livevalue < alertvalue:
        playsound('F:/ArunRS/Internship/Stock Market Web Scraping/ding-sound.mp3')
        sheet.range('D2').value= "Price Reached"
    else:
        difference= livevalue-alertvalue
        sheet.range('D2').value=difference'''

    
