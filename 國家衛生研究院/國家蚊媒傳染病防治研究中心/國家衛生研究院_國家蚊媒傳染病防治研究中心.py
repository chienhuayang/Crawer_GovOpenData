from itertools import zip_longest
from selenium.webdriver.common.by import By

from selenium import webdriver
import requests
from bs4 import BeautifulSoup
import time
from pandas import DataFrame
import openpyxl

#國家衛生研究院
#國家蚊媒傳染病防治研究中心

url = 'https://nmbdcrc.nhri.org.tw/investigators/researcher/'

enterprise = "國家衛生研究院"
groupname='國家蚊媒傳染病防治研究中心'

#excel製作
titles = ('全名','名字','中間名','姓氏','電子郵件','公司名稱','商務電話','學院','系所','行政單位一級')
wb = openpyxl.Workbook()
sheet_name = groupname
sheet = wb.create_sheet(sheet_name, 0)
sheet.append(titles)

r = requests.get(url)
soup = BeautifulSoup(r.text,'html.parser')

for a in soup.find_all('div',class_='wp-block-media-text__content'):
    if a.text:
        all_data = a.text
        datalist = all_data.split()
        fullname = datalist[0]
        names = list(fullname)
        email=datalist[-1]
        if all_data.find("分機")!=-1:
            tel_position = all_data.index("分機")
            tel = all_data[tel_position+3:tel_position+8]
            print(tel)
            phone = "06-700-0123#"+tel
        print(fullname)
        ap_list = [fullname,names[0],names[1],names[2] ,email, enterprise, phone,"","",groupname]
        sheet.append(ap_list)

filename='聯絡人_'+enterprise+"_"+groupname+'_工作人員'+'.xlsx'
wb.save(filename)

        