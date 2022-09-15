from selenium import webdriver
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import requests
import time
import openpyxl

r = requests.get('https://www.tesri.gov.tw/A5_0')
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
phonenum='(049)2761331'
page = driver.get('https://www.tesri.gov.tw/A5_0')

for i in driver.find_elements(By.XPATH,'//div[@class="col-md-4"]//div//a'):
    if i.get_attribute('href'):
        groupurl.append(i.get_attribute('href'))
print(groupurl)        

for _ in range(len(groupurl)):
    driver.get(groupurl[_])
    if driver.find_element(By.XPATH,'//div[@class="cont_newstitle"]').text: groupname=driver.find_element(By.XPATH,'//div[@class="cont_newstitle"]').text
    for a in driver.find_elements(By.XPATH,'//tbody//tr//td'):
        if a.text :
            data = a.text
            data= data.replace(' ','')
            if len(a.text)<10 and (data.find('1')!=-1 or data.find('9')!=-1 or data.find('8')!=-1 or data.find('7')!=-1 or data.find('6')!=-1 or data.find('4')!=-1 or data.find('5')!=-1 or data.find('2')!=-1 or data.find('3')!=-1 and data.find('@')==-1):
                phone = phonenum+'#'+data
                print(phone)
            elif data.find('@')!=-1 and len(a.text)<50:
                email = data
                print(email)
            elif len(data)<5 :
                name = data
                print(name)
            
            if name !='' and phone!='' and email !='':
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