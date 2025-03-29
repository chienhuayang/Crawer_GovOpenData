from selenium import webdriver
from selenium.webdriver.common.by import By
import requests
from bs4 import BeautifulSoup
import time
from pandas import DataFrame
import openpyxl
import os

#國家衛生研究院
#神經及精神醫學研究中心


enterprise = '國家衛生研究院'
groupname='免疫醫學研究中心'
filepath = ''+enterprise
name=''
email=''
phone=''


if os.path.isfile(filepath) !=-1:
    filepath = 'C:\\Users\\智域國際_楊倩華\\Documents\\202208\\爬蟲_聯絡人\\'+enterprise+'\\'+groupname
    if os.path.isfile(filepath)!=-1: 
        os.mkdir(filepath)
        print('file will be save at '+filepath)
    else: print('file will be save at '+filepath)
else:
    os.mkdir(filepath)
    filepath = 'C:\\Users\\智域國際_楊倩華\\Documents\\202208\\爬蟲_聯絡人\\'+enterprise+'\\'+groupname
    os.mkdir(filepath)
    print('file will be save at '+filepath)


    

ap_list=[]
pagelist=['https://im.nhri.org.tw/index.php/staff-directory/']

#excel
title = ('全名','名字','中間名','姓氏','電子郵件','公司名稱','商務電話','學院','系所','行政單位一級')
wb = openpyxl.Workbook()
sheet_name = groupname
sheet = wb.create_sheet(sheet_name,0)
sheet.append(title)

driver = webdriver.Chrome()

phonenum = '037-206-166'
for a in range(len(pagelist)):
    page = driver.get(pagelist[a])
    for i in driver.find_elements(By.XPATH,"//tbody//tr//td"): 
        #i只有一格資料
        #['莊宗顯'] #['研究員暨代理主任'] #['分機37611'] #['thchuang@nhri.edu.tw']
        if i.text:
            data = i.text
            if data.find('研究')!=-1:print("none")
            elif data.find('分機')!=-1:
                data = data.replace("分機",'#')
                phone = phonenum+data
            elif data.find('@')!=-1:
                email = data
            else:
                name = data
        if name !='' and email!='' and phone!='':
            names = list(name)
            if len(names)<3:ap_list = [name,names[1],"",names[0] ,email, enterprise, phone,"","",groupname]
            else:ap_list = [name,names[1]+names[2],"",names[0] ,email, enterprise, phone,"","",groupname]
            name=''
            email=''
            phone=''
            sheet.append(ap_list)
filename=filepath+'\\聯絡人_'+enterprise+groupname+'.xlsx'
wb.save(filename)



