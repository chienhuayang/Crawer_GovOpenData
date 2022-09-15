from fileinput import filename
from turtle import title
from selenium import webdriver
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import time
import requests
import openpyxl

r = requests.get('https://www.iner.gov.tw/%E8%81%AF%E7%B5%A1%E6%88%91%E5%80%91--7.html')
soup = BeautifulSoup(r.text,'html.parser') 
E_title = soup.title.text
E_title = E_title.split() 
enterprise= E_title[-1]
phone=''
name=''
email=''
groupname=''
groupnamelist=[]

#excel
titles = ('全名','名字','中間名','姓氏','電子郵件','公司名稱','商務電話','學院','系所','行政單位一級')
wb = openpyxl.Workbook()
sheet = wb.create_sheet(enterprise,0)
sheet.append(titles)

driver = webdriver.Chrome()
i=0
y=0
phonenum = '(03)471-1400'
page = driver.get('https://www.iner.gov.tw/%E8%81%AF%E7%B5%A1%E6%88%91%E5%80%91--7.html')

for a in driver.find_elements(By.XPATH,"//div//p"):
    groupname = a.text
    groupnamelist.append(groupname)
    print(groupnamelist)


for b in driver.find_elements(By.XPATH,"//tbody//tr//td"):
    if b.text and len(b.text)<15:
        data = b.text
        all = data.split()
        if i==0: 
            name = all[-1]
            i+=1
        elif i==1:
            tel = all[-1]
            phone = phonenum+'#'+tel
            i+=1
        elif i==2:
            email=all[-1]
            email = email+'@iner.gov.tw'
        
        if name !='' and phone!='' and email!='' :
                names = list(name)
                if len(names)<3:ap_list = [name,names[1],"",names[0] ,email, enterprise, phone,"","",groupnamelist[y]]
                else:ap_list = [name,names[1]+names[2],"",names[0] ,email, enterprise, phone,"","",groupnamelist[y]]
                name=''
                phone=''
                email=''
                groupname=''
                i=0
                y+=1
                print(ap_list)
                sheet.append(ap_list)


       
filename = enterprise+'.xlsx'
wb.save(filename)

