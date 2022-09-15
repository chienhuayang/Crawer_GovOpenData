from selenium import webdriver
from selenium.webdriver.common.by import By
import requests
from bs4 import BeautifulSoup
import time
from pandas import DataFrame
import openpyxl
import os

enterprise = '國家衛生研究院'
groupname='群體健康科學研究所'
filepath = 'C:\\Users\\智域國際_楊倩華\\Documents\\202208\\爬蟲_聯絡人\\'+enterprise
name=''
email=''
phone=''


if os.path.isfile(filepath) !=-1:
    filepath = filepath+'\\'+groupname
    
    if os.path.isfile(filepath)==-1: 
        os.mkdir(filepath)
        print('file will be save at '+filepath)
    else: print('file will be save at '+filepath)
else:
    os.mkdir(filepath)
    filepath = filepath+enterprise+'\\'+groupname
    os.mkdir(filepath)
    print('file will be save at '+filepath)

#excel
title = ('全名','名字','中間名','姓氏','電子郵件','公司名稱','商務電話','學院','系所','行政單位一級')
wb = openpyxl.Workbook()
sheet_name = groupname
sheet = wb.create_sheet(sheet_name,0)
sheet.append(title)
    

ap_list=[]
pagelist=['https://ph.nhri.org.tw/zhtw/%e7%a0%94%e7%a9%b6%e6%88%90%e5%93%a1/']

driver = webdriver.Chrome()
page = driver.get(pagelist[0])


for i in driver.find_elements(By.XPATH,"//div[@class='pi_data']"):
    if i.text:
        data = i.text
        #data = data.replace('/n',' ')
        data = data.replace('：',"\n")
        data = data.replace('/',"\n")
        data = data.replace("'","")
        y_data = data.split('\n')
        print(y_data)
        #y_data=['陳美惠 ', ' Mei-Huei Chen', '主治醫師 ', ' Attending physician', 'Chenmh@nhri.edu.tw', 'Learn about me…', 'Institutional Repository']

        name = y_data[0]
        names = y_data[1].split()
        for a in y_data:
            if a.find('@')!=-1:
                email = a
        

        if len(names)<3:ap_list = [name,names[0],"",names[1] ,email, enterprise, phone,"","",groupname]
        else:ap_list = [name,names[1],names[1],names[2] ,email, enterprise, phone,"","",groupname]
        sheet.append(ap_list)
filename=filepath+'\\聯絡人_'+enterprise+groupname+'.xlsx'
wb.save(filename)