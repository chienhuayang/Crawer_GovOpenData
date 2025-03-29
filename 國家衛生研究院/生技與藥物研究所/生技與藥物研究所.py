from selenium import webdriver
from selenium.webdriver.common.by import By
import requests
from bs4 import BeautifulSoup
import time
from pandas import DataFrame
import openpyxl
import os



enterprise = '國家衛生研究院'
groupname='生技與藥物研究所'
filepath = ''
pagelist=['https://ibpr.nhri.org.tw/zhtw/index.php/investigators-2/','https://ibpr.nhri.org.tw/zhtw/index.php/research-staff/','https://ibpr.nhri.org.tw/zhtw/index.php/administrative-staff/']
name=''
email=[]
y=0
phone=''




filename='聯絡人_'+enterprise+groupname+'.xlsx'
    

ap_list=[]


#excel
title = ('全名','名字','中間名','姓氏','電子郵件','公司名稱','商務電話','學院','系所','行政單位一級')
wb = openpyxl.Workbook()
sheet_name = groupname
sheet = wb.create_sheet(sheet_name,0)
sheet.append(title)


driver = webdriver.Chrome()

phonenum = '037-206166'


for a in range(len(pagelist)):
    page = driver.get(pagelist[a])
    y=0
    email=[]
    for lnks in driver.find_elements(By.XPATH,"//tbody//tr//td//a"):
        lnk = lnks.get_attribute('href')
        lnk = lnk.replace('mailto:','')
        print(lnk)
        if lnk.find('@')!=-1 and lnk.find('https')==-1 and lnk.find('surprise1986@nhri.org.tw')==-1: email.append(lnk)

    for b in driver.find_elements(By.XPATH,"//tbody//tr//td"):
            
        if b.text:
            data = b.text
            #print(data+"====")
            if data.find('研究')!=-1 or data.find('E-mail')!=-1 or data.find('院')!=-1 or data.find('分機')!=-1 or data.find('職稱')!=-1 or data.find('姓名')!=-1 or data.find('計畫')!=-1 or data.find('秘書')!=-1 or data.find('主任')!=-1 or data.find('人員')!=-1 or data.find('實驗室')!=-1 or data.find('助理')!=-1 or data.find('師')!=-1:data=''
            if data.find('3')!=-1 and data.find('@')==-1:
                phone = phonenum+data
                if len(phone)>18:
                    phone = data
            
            else:
                name = data
                
            
            if name !='' and phone!='' and name !=' ':
                names = list(name)
                if len(names)<3:ap_list = [name,names[1],"",names[0] ,email[y], enterprise, phone,"","",groupname]
                else:ap_list = [name,names[1]+names[2],"",names[0] ,email[y], enterprise, phone,"","",groupname]
                name=''
                phone=''
                y+=1
                print(ap_list)
                sheet.append(ap_list)

wb.save(filename)

