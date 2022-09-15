from selenium.webdriver.common.by import By

from selenium import webdriver
import requests
from bs4 import BeautifulSoup
import time
from pandas import DataFrame
import openpyxl

#!/usr/bin/python
# -*- coding: cp65001 -*-
#國家倫理科學研究中心物理組

#driver = webdriver.Chrome()
url = 'https://phys.ncts.ntu.edu.tw/'
r = requests.get("https://phys.ncts.ntu.edu.tw/en/people/director_office",verify=False)
soup = BeautifulSoup(r.text,'html.parser') #將網頁資料以
enterprise_tag = soup.title #找公司名
enterprise_name = enterprise_tag.string
enterprise_name = enterprise_name.replace("\t","")
enterprise_name = enterprise_name.replace("\r\n","")
etps_split = enterprise_name.split()
enterprise = etps_split[3]
groupname=etps_split[5]
print(enterprise,groupname)
groupurl_list=[]
perlist = []



#excel製作
titles = ('全名','名字','中間名','姓氏','電子郵件','公司名稱','商務電話','學院','系所','行政單位一級')
wb = openpyxl.Workbook()
sheet_name = etps_split[3]+etps_split[4]+etps_split[5]
sheet = wb.create_sheet(sheet_name, 0)
sheet.append(titles)

groupurl_list=['https://phys.ncts.ntu.edu.tw/en/people/center_scientists','https://phys.ncts.ntu.edu.tw/en/people/research_staff','https://phys.ncts.ntu.edu.tw/en/people/member4']    


for i in range(len(groupurl_list)) :
    
    #group_page = driver.get(groupurl_list[i])
    r = requests.get(groupurl_list[i],verify=False)
    soup = BeautifulSoup(r.text,'html.parser')
    #a=soup.findAll('a',class_="i-member-link",href=True)
    #print(a)

    for a in soup.find_all('a',class_="i-member-link", href=True): 
         if a.text: 
            phone = ""
            email = ""
            per_url = url+a['href'] #取得個人網頁
            r2 = requests.get(per_url)
            per_page = BeautifulSoup(r2.text,'html.parser')
            name = per_page.find('td',class_="member-data-value-name").string
            name = name.replace(",","")
            name = name.replace('-'," ")
            print(name)
            namelist = name.split()
            
            if per_page.find('td',class_="member-data-value-email"):email = per_page.find('td',class_="member-data-value-email").string
            if per_page.find('td',class_="member-data-value-office-tel"): 
                phone = per_page.find('td',class_="member-data-value-office-tel").string
                phone.replace('+886-','0')

            if len(namelist)==4:  ap_list = [namelist[-1],namelist[1],namelist[2],namelist[0] ,email, enterprise, phone,"","",groupname]
            elif len(namelist)<4: ap_list = [namelist[-1],namelist[1],'',namelist[0] ,email, enterprise, phone,"","",groupname]

         filename='聯絡人_'+enterprise+"_"+groupname+'.xlsx'
         sheet.append(ap_list)
         wb.save(filename)
         time.sleep(3)
                
