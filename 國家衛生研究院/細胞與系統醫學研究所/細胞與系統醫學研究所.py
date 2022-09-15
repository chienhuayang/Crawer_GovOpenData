from selenium import webdriver
from selenium.webdriver.common.by import By
import requests
from bs4 import BeautifulSoup
import time
from pandas import DataFrame
import openpyxl
import os

enterprise = '國家衛生研究院'
groupname='細胞與系統醫學研究所'
filepath = 'C:\\Users\\智域國際_楊倩華\\Documents\\202208\\爬蟲_聯絡人\\'+enterprise
name=''
email=''
phone=''
phonenum = ' (037)206166'

if os.path.isfile(filepath) !=-1:
    filepath = filepath+'\\'+groupname
    
    if os.path.isfile(filepath)==-1: 
        os.mkdir(filepath)
        print('file will be save at '+filepath)
    elif os.path.isfile(filepath)!=-1: print('file will be save at '+filepath)
else:
    os.mkdir(filepath)
    filepath = filepath+'\\'+groupname
    os.mkdir(filepath)
    print('file will be save at '+filepath)

#excel
title = ('全名','名字','中間名','姓氏','電子郵件','公司名稱','商務電話','學院','系所','行政單位一級')
wb = openpyxl.Workbook()
sheet_name = groupname
sheet = wb.create_sheet(sheet_name,0)
sheet.append(title)
    

ap_list=[]
pagelist=['https://cs.nhri.org.tw/tw/investigators/','https://cs.nhri.org.tw/tw/honorary-investigators/','https://cs.nhri.org.tw/tw/joint-appointed-adjunct-investigator/']

driver = webdriver.Chrome()


for a in range(len(pagelist)):
    page = driver.get(pagelist[a])
    for i in driver.find_elements(By.XPATH,"//tbody//tr//td"):
        if i.text:
            data = i.text
            #data = data.replace('/n',' ')
            data = data.replace('：',"\n")
            data = data.replace('/',"\n")
            data = data.replace("'","")
            data = data.replace(" ","\n")
            y_data = data.split('\n')
            print(y_data)

            name = y_data[0]
            
            for y in y_data:
                if y.find('03')!=-1 or y.find('02')!=-1 and y.find('@')==-1:
                    y=y.replace('Tel:','')
                    phone = y
                elif y.find('@')!=-1:
                    y=y.replace('Email:','')
                    email = y
            if name !='' and email!='' and phone!='':
                names=list(name)
                name = names[0]+names[1]
                if len(names)<3:ap_list = [name,names[0],"",names[1] ,email, enterprise, phone,"","",groupname]
                elif len(names)>=3:ap_list = [name+names[2],names[1]+names[2],"",names[0] ,email, enterprise, phone,"","",groupname]
                sheet.append(ap_list)
                name,email,phone='','',''
#行政人員
page = driver.get('https://cs.nhri.org.tw/tw/staff-directory/')
for i in driver.find_elements(By.XPATH,"//tbody//tr//td"): 
        if i.text:
            data = i.text
            if data.find('行政')!=-1:print("none")
            elif data.find('3')!=-1 and data.find('@')==-1:
                #data = data.replace("分機",'#')
                phone = phonenum+data
                #print(data)
            elif data.find('@')!=-1:
                email = data
                #print(data)
            else:
                name = data
                #print(data)
        if name !='' and email!='' and phone!='':
            names = list(name)
            if len(names)<3:ap_list = [name,names[1],"",names[0] ,email, enterprise, phone,"","",groupname]
            elif len(names)>=3:ap_list = [name,names[1]+names[2],"",names[0] ,email, enterprise, phone,"","",groupname]
            name=''
            email=''
            phone=''
            sheet.append(ap_list)
#             name = y_data[0]
#             names = y_data[1].split()
#             for a in y_data:
#                 if a.find('@')!=-1:
#                     email = a
        

            
filename='聯絡人_'+enterprise+groupname+'.xlsx'
wb.save(filename)