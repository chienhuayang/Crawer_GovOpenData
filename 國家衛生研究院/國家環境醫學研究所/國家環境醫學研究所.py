from selenium import webdriver
from selenium.webdriver.common.by import By
import requests
from bs4 import BeautifulSoup
import time
from pandas import DataFrame
import openpyxl
import os



enterprise = '國家衛生研究院'
groupname='國家環境醫學研究所'
filepath =''
name=''
email=''
phone=''


if os.path.isfile(filepath) !=-1:
    filepath =''
    if os.path.isfile(filepath)==-1: 
        os.mkdir(filepath)
        print('file will be save at '+filepath)
    else: print('file will be save at '+filepath)
else:
    os.mkdir(filepath)
    filepath = ''
    os.mkdir(filepath)
    print('file will be save at '+filepath)


    

ap_list=[]
pagelist=['https://niehs.nhri.edu.tw/%e7%a0%94%e7%a9%b6%e6%88%90%e5%93%a1']

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
    for i in driver.find_elements(By.XPATH,"//tbody//tr//td"): 
        #i只有一格資料
        #['莊宗顯'] #['研究員暨代理主任'] #['分機37611'] #['thchuang@nhri.edu.tw']
        if i.text:
            data = i.text
            #data = data.replace('/n',' ')
            data = data.replace('：',"\n")
            data = data.replace(':',"\n")
            data = data.replace("'","")
            y_data = data.split('\n')
            print(y_data)
            #y_data=['陳保中特聘研究員', 'Email', ' pchen@nhri.edu.tw', '分機', '36500']
            name = y_data[0]
            name = name.replace('副研究員','')
            name = name.replace('助研究員','')
            name = name.replace('名譽研究員','')
            name = name.replace('特聘','')
            name = name.replace('合聘','')
            name = name.replace('研究員','')
            name = name.replace('主治醫師','')
            name = name.replace('兼任','')
            
                    
            phone = phonenum +"#"+y_data[-1] 
            if len(phone)>18 and len(phone)<=12:
                phone = y_data[-1]
            email = y_data[2]
            names = list(name)
            if len(names)<3:ap_list = [name,names[1],"",names[0] ,email, enterprise, phone,"","",groupname]
            else:ap_list = [name,names[1]+names[2],"",names[0] ,email, enterprise, phone,"","",groupname]

            print(y_data)
            sheet.append(ap_list)
filename=filepath+'\\聯絡人_'+enterprise+groupname+'.xlsx'
wb.save(filename)
