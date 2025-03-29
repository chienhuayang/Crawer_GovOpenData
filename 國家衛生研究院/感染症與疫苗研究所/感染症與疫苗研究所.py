
from selenium import webdriver
from selenium.webdriver.common.by import By
import requests
from bs4 import BeautifulSoup
import time
from pandas import DataFrame
import openpyxl
import os



enterprise = '國家衛生研究院'
groupname='感染症與疫苗研究所'
filepath = ''
#pagelist=['https://ibpr.nhri.org.tw/zhtw/index.php/investigators-2/','https://ibpr.nhri.org.tw/zhtw/index.php/research-staff/','https://ibpr.nhri.org.tw/zhtw/index.php/administrative-staff/']
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


page = driver.get('https://iv.nhri.org.tw/zhtw/staff-directory/')
for i in driver.find_elements(By.XPATH,"//tbody//tr//td"): 
        if i.text:
            data = i.text
            data = data.replace("室主任",'')
            data = data.replace("(兼任副院長)",'')
            if data.find('室')!=-1 or data.find('品質')!=-1 or data.find('業務')!=-1 or data.find('廠務')!=-1 or data.find('開發')!=-1 or data.find('行政')!=-1 or data.find('機')!=-1 or data.find('通訊錄')!=-1 or data.find('郵件')!=-1 or data.find('生物')!=-1 or data.find('大學')!=-1 or data.find('管')!=-1 or data.find('名')!=-1 or data.find('電話')!=-1 or data.find('所長')!=-1 or data.find('研究')!=-1 or data.find('E-mail')!=-1 or data.find('院')!=-1 or data.find('分機')!=-1 or data.find('職')!=-1 or data.find('姓名')!=-1 or data.find('計畫')!=-1 or data.find('秘書')!=-1 or data.find('人員')!=-1 or data.find('實驗室')!=-1 or data.find('助理')!=-1 or data.find('師')!=-1:data=''
            elif data.find('6')!=-1 or data.find('3')!=-1 and data.find('@')==-1:
                #data = data.replace("分機",'#')
                phone = phonenum+"#"+data
                print(phone)
            elif data.find('@')!=-1:
                email = data
                print(email)
            else:
                rname = data
                rname = rname.replace("執行長",'')
                rname = rname.replace("科長",'')
                rname = rname.replace("代理",'')
                rname = rname.replace("主任",'')
                aname = rname.split()
                if aname!=[]:
                    if len(aname[-1])<=3 :
                        name = aname[-1]
                        print('name:',name)
            if data.find('司徒惠康')!=-1:
                name = data
            
        if name !='' and email!='' and phone!='':
            names = list(name)
            if len(names)<3:ap_list = [name,names[1],"",names[0] ,email, enterprise, phone,"","",groupname]
            elif len(names)>=3:ap_list = [name,names[1]+names[2],"",names[0] ,email, enterprise, phone,"","",groupname]
            name=''
            email=''
            phone=''
            sheet.append(ap_list)

        

            
filename='聯絡人_'+enterprise+groupname+'.xlsx'
wb.save(filename)
