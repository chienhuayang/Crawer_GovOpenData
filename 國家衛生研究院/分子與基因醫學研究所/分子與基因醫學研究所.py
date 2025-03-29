from selenium import webdriver
from selenium.webdriver.common.by import By
import requests
from bs4 import BeautifulSoup
import time
from pandas import DataFrame
import openpyxl
import os



enterprise = '國家衛生研究院'
groupname='分子與基因醫學研究所'
filepath = ''
name=''
email=''
phone=''

print("hello!")
if os.path.isfile(filepath) !=-1:
    filepath = ''
    if os.path.isfile(filepath)!=-1: 
        
        print('file will be save at '+filepath)
    else: 
        os.mkdir(filepath)
        print('file will be save at '+filepath)
else:
    os.mkdir(filepath)
    filepath = ''
    os.mkdir(filepath)
    print('file will be save at '+filepath)

filename=filepath+'\\聯絡人_'+enterprise+groupname+'.xlsx'
    

ap_list=[]
pagelist=['https://mg.nhri.org.tw/tw/staff-directory/']

#excel
title = ('全名','名字','中間名','姓氏','電子郵件','公司名稱','商務電話','學院','系所','行政單位一級')
wb = openpyxl.Workbook()
sheet_name = groupname
sheet = wb.create_sheet(sheet_name,0)
sheet.append(title)
wb.save(filename)

driver = webdriver.Chrome()

phonenum = '037-206166'
page = driver.get('https://mg.nhri.org.tw/tw/staff-directory/')

for b in driver.find_elements(By.XPATH,"//tbody//tr//td"):
    if b.text:
        data = b.text
        
        if data.find('研究')!=-1 or data.find('人員')!=-1 or data.find('實驗室')!=-1 or data.find('助理')!=-1 or data.find('師')!=-1:data=''
        if data.find('3')!=-1 and data.find('@')==-1:
             phone = phonenum+data
             if len(phone)>18:
                phone = data
        elif data.find('@')!=-1:
            email = data
            names = list(name)
        else:
            name = data
        if name !='' and email!='' and phone!='' and name !=' ':
            names = list(name)
            if len(names)<3:ap_list = [name,names[1],"",names[0] ,email, enterprise, phone,"","",groupname]
            else:ap_list = [name,names[1]+names[2],"",names[0] ,email, enterprise, phone,"","",groupname]
            name=''
            email=''
            phone=''
            print(ap_list)
            sheet.append(ap_list)

wb.save(filename)

