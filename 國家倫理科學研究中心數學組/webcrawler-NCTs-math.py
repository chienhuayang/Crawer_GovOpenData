from itertools import zip_longest
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
url = 'https://ncts.ntu.edu.tw/'
r = requests.get("https://ncts.ntu.edu.tw",verify=False)
soup = BeautifulSoup(r.text,'html.parser') #將網頁資料以
enterprise_tag = soup.title #找公司名
enterprise_name = enterprise_tag.string
enterprise_name = enterprise_name.replace("\t","")
enterprise_name = enterprise_name.replace("\r\n","")
etps_split = enterprise_name.split()
enterprise = "國家倫理科學研究中心"
groupname='數學組'
#print(enterprise,groupname)
groupurl_list=[]
k=0

#excel製作
titles = ('全名','名字','中間名','姓氏','電子郵件','公司名稱','商務電話','學院','系所','行政單位一級')
wb = openpyxl.Workbook()
sheet_name = "工作人員"
sheet = wb.create_sheet(sheet_name, 0)
sheet.append(titles)


groupurl_list=['https://ncts.ntu.edu.tw/people_list_9.php?bgid=9']    


for s in range(len(groupurl_list)) :
    #print('a')
    #group_page = driver.get(groupurl_list[i])
    r = requests.get(groupurl_list[s],verify=False)
    soup = BeautifulSoup(r.text,'html.parser')
    listdata = []
    listname=[]
    
    #a=soup.findAll('a',class_="i-member-link",href=True)
    #print(a)

    for a in soup.find_all('td',class_="peopleName"): 
        if a.text: 
            fullname = a.text
            #print('name:',fullname)
            listname.append(fullname)
            for b in soup.find_all('td',class_="peopleText1"):
                if b.text:
                    data = b.text
                    data=data.replace("\t",'')
                    data=data.replace(" ",'')
                    data=data.replace("\r\n",'')
                    listdata.append(data)  
    
    #print(listdata) 

    
    for i,ks in zip(range(len(listname)),range(len(listdata))):
        print(k)
        name = listname[i]
        name = name.replace('-',' ')
        name_split=name.split()
        if len(name_split)>4:
            ch_name = name_split[-2]+name_split[-1]
            names = [name_split[0],name_split[1],name_split[2]]
        elif len(name_split)==3:
            ch_name = name
            names = [name_split[0],name_split[1],name_split[2]]
        elif len(name_split)==4:
            ch_name = name_split[-2]+name_split[-1]
            names = [name_split[0]," ",name_split[1]]
        else:
            ch_name = name
            names = [name_split[0]," ",name_split[1]]
        
        email = listdata[k+1]
        all_d =  listdata[k+3]
        detail = all_d.split(",")
        phone = listdata[k+2]
        phone=phone.replace('+886-','0')
        school = detail[-1]
        email = email.replace("[AT]","@")
        #email = email.replace("AT","@")
        #email = email.replace("at","@")
        email = email.replace("m@h","math")
        ap_list = [ch_name,names[0],names[1],names[2] ,email, enterprise, phone,school,"",groupname]
        sheet.append(ap_list)
        k = k+4
 
        

         
filename='聯絡人_'+enterprise+"_"+groupname+'_工作人員'+'.xlsx'
wb.save(filename)
time.sleep(3)
                
