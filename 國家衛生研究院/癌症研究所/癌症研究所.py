from re import I
from tkinter import Y
from selenium.webdriver.common.by import By

from selenium import webdriver
import requests
from bs4 import BeautifulSoup
import time
from pandas import DataFrame
import openpyxl


#driver = webdriver.Chrome()
url = 'https://nicr.nhri.org.tw/pi/'
r = requests.get("https://nicr.nhri.org.tw/pi/",verify=False)
soup = BeautifulSoup(r.text,'html.parser') #將網頁資料以
enterprise='國家衛生研究院'
groupname = '癌症研究所'
per_url=[]
driver = webdriver.Chrome()

#excel製作
titles = ('全名','名字','中間名','姓氏','電子郵件','公司名稱','商務電話','學院','系所','行政單位一級')
wb = openpyxl.Workbook()
sheet_name = groupname
sheet = wb.create_sheet(sheet_name, 0)
sheet.append(titles)

#groupurl_list=['https://phys.ncts.ntu.edu.tw/en/people/center_scientists','https://phys.ncts.ntu.edu.tw/en/people/research_staff','https://phys.ncts.ntu.edu.tw/en/people/member4']    

page = driver.get(url)

for i in driver.find_elements(By.XPATH,"//tbody//tr//td//a"):
    lnk = i.get_attribute('href')
    per_url.append(lnk)
for u in range(len(per_url)):
    per = driver.get(per_url[u])
    for a in driver.find_elements(By.XPATH,"//tbody//tr//td[@class='word_1']"):
        data = a.text
        data = data.replace('：',"\n")
        data = data.replace(':',"\n")
        data = data.replace("'","")
        y_data = data.split('\n')
        #y_data=['夏興國', '副研究員', '癌症研究所', '電話', '037-246166 ext. 31707 / 31712', 'E-mail', 'davidssg@nhri.edu.tw']
        #print(y_data)

        name = y_data[0]
        for y in y_data:
            if y.find('@')!=-1:
                email=y
            elif y.find('3')!=-1 and y.find('@')==-1:
                phone = y
        names = list(name)
        if len(names)<3:ap_list = [name,names[1],"",names[0] ,email, enterprise, phone,"","",groupname]
        elif len(names)>=3:ap_list = [name,names[1]+names[2],"",names[0] ,email, enterprise, phone,"","",groupname]
        print(ap_list)
        sheet.append(ap_list)
filename = filename='聯絡人_'+enterprise+groupname+'.xlsx'
wb.save(filename)

    
        

        