from selenium import webdriver
from selenium.webdriver.common.by import By
import requests
from bs4 import BeautifulSoup
import time
from pandas import DataFrame
import openpyxl
import os



enterprise = '國家衛生研究院'
groupname='生醫工程與奈米醫學研究所'
filepath = 'C:\\Users\\智域國際_楊倩華\\Documents\\202208\\爬蟲_聯絡人\\'+enterprise
name=''
email=''
phone=''
phonenum = '037-206166'

if os.path.isfile(filepath) !=-1:
    filepath = 'C:\\Users\\智域國際_楊倩華\\Documents\\202208\\爬蟲_聯絡人\\'+enterprise+'\\'+groupname
    if os.path.isfile(filepath)==-1: 
        os.mkdir(filepath)
        print('file will be save at '+filepath)
    elif os.path.isfile(filepath)!=-1: print('file will be save at '+filepath)
else:
    os.mkdir(filepath)
    filepath = 'C:\\Users\\智域國際_楊倩華\\Documents\\202208\\爬蟲_聯絡人\\'+enterprise+'\\'+groupname
    os.mkdir(filepath)
    print('file will be save at '+filepath)


    

ap_list=[]
pagelist=['http://iben.nhri.org.tw/staffs.php']

#excel
title = ('全名','名字','中間名','姓氏','電子郵件','公司名稱','商務電話','學院','系所','行政單位一級')
wb = openpyxl.Workbook()
sheet_name = groupname
sheet = wb.create_sheet(sheet_name,0)
sheet.append(title)

driver = webdriver.Chrome()


for a in range(len(pagelist)):
    page = driver.get(pagelist[a])
    for i in driver.find_elements(By.XPATH,"//div[@class='caption']"): 
        if i.text:
            data = i.text
            #data = data.replace('/n',' ')
            data = data.replace('：',"\n")
            data = data.replace(':',"\n")
            data = data.replace("'","")
            data = data.replace(" ","\n")

            y_data = data.split('\n')
            print(y_data)
            #y_data=['董國忠', '博士', '副研究員', '辦公室分機', '37135', '實驗室分機', '37155', 'gcdong@nhri.edu.tw']
            name = y_data[0]
                       
                    
            phone = phonenum +"#"+y_data[4] 
            # if len(phone)>18 and len(phone)<=12:
            #     phone = y_data[-1]
            email = y_data[7]
            names = list(name)
            if len(names)<3:ap_list = [name,names[1],"",names[0] ,email, enterprise, phone,"","",groupname]
            else:ap_list = [name,names[1]+names[2],"",names[0] ,email, enterprise, phone,"","",groupname]

            
            sheet.append(ap_list)
filename=filepath+'\\聯絡人_'+enterprise+groupname+'.xlsx'
wb.save(filename)