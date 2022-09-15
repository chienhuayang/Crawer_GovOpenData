from selenium import webdriver
from selenium.webdriver.common.by import By
import requests
from bs4 import BeautifulSoup
import time
from pandas import DataFrame
import openpyxl

#國家衛生研究院
#神經及精神醫學研究中心

url = 'https://np.nhri.org.tw/full-time-investigator/'

enterprise = '國家衛生研究院'
groupname='神經及精神醫學研究中心'
ap_list=[]

#excel
title = ('全名','名字','中間名','姓氏','電子郵件','公司名稱','商務電話','學院','系所','行政單位一級')
wb = openpyxl.Workbook()
sheet_name = groupname
sheet = wb.create_sheet(sheet_name,0)
sheet.append(title)

driver = webdriver.Chrome()
page = driver.get(url)
phonenum = '037-206-166'
for i in driver.find_elements(By.XPATH,"//tbody//tr//td"):
    data = i.text
    #data = data.replace('/n',' ')
    data = data.replace('：',"\n")
    data = data.replace("'","")
    
    y_data = data.split('\n')
    #y_data=["姓名'", "'陳為堅", "職稱'", "'主任/特聘研究員", "分機'", "'36700", "郵件'", "' wjchen @nhri.edu.tw"]
    name = y_data[1]
    phone = phonenum +"#"+y_data[5] 
    email = y_data[-1]
    names = list(name)
    ap_list = [name,names[1]+names[2],"",names[0] ,email, enterprise, phone,"","",groupname]
    print(y_data)
    sheet.append(ap_list)
filename='聯絡人_'+enterprise+groupname+'.xlsx'
wb.save(filename)



