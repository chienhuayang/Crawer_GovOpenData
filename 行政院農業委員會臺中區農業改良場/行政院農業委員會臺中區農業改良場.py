from selenium import webdriver
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import requests
import time
import openpyxl

r = requests.get('https://www.tdais.gov.tw/index.php')
soup = BeautifulSoup(r.text,'html.parser') 
E_title = soup.title.text
E_title = E_title.split() 
enterprise= E_title[-1]
print(enterprise)
phone=''
name=''
email=''
groupname=''
groupnamelist=[]
groupurl=[]

#excel
titles = ('全名','名字','中間名','姓氏','電子郵件','公司名稱','商務電話','學院','系所','行政單位一級')
wb = openpyxl.Workbook()
sheet = wb.create_sheet(enterprise,0)
sheet.append(titles)
driver = webdriver.Chrome()
phonenum=''
page = driver.get('https://www.tdais.gov.tw/ws.php?id=5484')

for b in driver.find_elements(By.XPATH,'//tbody//tr//td//strong'):
    if b.text:
        groupnamelist.append(b.text)
print(groupnamelist)    
i=0
for a in driver.find_elements(By.XPATH,'//tbody//tr//td'):
        if a.text :
            data = a.text
            data= data.replace(' ','')
            if data.find("(由")!=-1:
                datalist=data.split("(")
                data=datalist[0]
            if (data.find('1')!=-1 or data.find('9')!=-1 or data.find('8')!=-1 or data.find('7')!=-1 or data.find('6')!=-1 or data.find('4')!=-1 or data.find('5')!=-1 or data.find('2')!=-1 or data.find('3')!=-1 and data.find('@')==-1):
                phone =data
                print(phone)
            elif data.find('@')!=-1 and len(a.text)<50:
                email = data
                print(email)
            
            elif len(data)<=4 and data.find("技佐")==-1 and data.find("書記")==-1 and data.find("主任")==-1 and data.find("研究")==-1 and data.find("場")==-1 and data.find("員")==-1 and data.find("室")==-1:
                name = data
                print(name)
            
            if i<len(groupnamelist) and data == groupnamelist[i]:
                groupname=groupnamelist[i]
                i+=1
            
            
            
            if name !='' and phone!='':
                names = list(name)
                if len(names)<3:ap_list = [name,names[1],"",names[0] ,email, enterprise, phone,"","",groupname]
                else:ap_list = [name,names[1]+names[2],"",names[0] ,email, enterprise, phone,"","",groupname]
                name=''
                phone=''
                email=''
                print(ap_list)
                sheet.append(ap_list)

filename='聯絡人_'+enterprise+'.xlsx'            
wb.save(filename)      