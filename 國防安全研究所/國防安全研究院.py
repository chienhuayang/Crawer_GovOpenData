from selenium import webdriver
from selenium.webdriver.common.by import By
import requests
from bs4 import BeautifulSoup
import time
import openpyxl

enterprise ='國防安全研究院'
#filepath = 'C:\Users\智域國際_楊倩華\OneDrive - 智域國際股份有限公司\文件\202208\爬蟲_聯絡人'
filename='聯絡人_'+enterprise+'.xlsx'
driver = webdriver.Chrome()
ap_list=[]
url=[]
#excel
title = ('全名','名字','中間名','姓氏','電子郵件','公司名稱','商務電話','學院','系所','行政單位一級')
wb = openpyxl.Workbook()
sheetname = enterprise
sheet = wb.create_sheet(sheetname,0)
sheet.append(title)
for i in range(4):
    driver.get('https://indsr.org.tw/researchinlist?uid=2&resid='+str(i+1))
    url=[]
    for a in driver.find_elements(By.XPATH,'//div//a[@class="figure flex flex-items-center"]'):
         if a.get_attribute('href'):
             perurl=a.get_attribute('href')
             url.append(perurl)
    for c in driver.find_elements(By.XPATH,'//div//a[@class="col-md-6"]'):
        if c.get_attribute('href'):
            print('find')
            perurl=c.get_attribute('href')
            #print(perurl)
            url.append(perurl) 
            print(url)       
    for b in range(len(url)) :
            driver.get(url[b])    
            if driver.find_element(By.XPATH,'//div[@class="firstTop"]'):
                data = driver.find_element(By.XPATH,'//div[@class="firstTop"]').text
                data.replace('\n','')
                alldata = data.split(' ')
                name = alldata[0]
                print(name)
            if driver.find_elements(By.XPATH,"//div[@class='way']//a"):
                lnks = driver.find_element(By.XPATH,"//div[@class='way']//a")
                lnk = lnks.get_attribute('href')
                lnk = lnk.replace('mailto:','')
                email = lnk
                print(email)
            if driver.find_elements(By.XPATH,"//div[@class='phone']"):
                phone = driver.find_element(By.XPATH,"//div[@class='phone']").text
                phone = phone.replace('\n','')
                phone = phone.replace('分機','#')
                print(phone)
            if driver.find_elements(By.TAG_NAME,"h2"):
                groupname=driver.find_element(By.TAG_NAME,"h2").text
                print(groupname)
            if name !='' and email!='' and phone!='' and groupname!='' :
                names = list(name)
                if len(names)<3:ap_list = [name,names[1],"",names[0] ,email, enterprise, phone,"","",groupname]
                elif len(names)>=3:ap_list = [name,names[1]+names[2],"",names[0] ,email, enterprise, phone,"","",groupname]
                sheet.append(ap_list)
            name=''
            email=''
            phone=''
            groupname=''
            time.sleep(2)
wb.save(filename)
            
            
            