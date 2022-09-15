from selenium.webdriver.common.by import By

from selenium import webdriver
import requests
from bs4 import BeautifulSoup
import time
from pandas import DataFrame
import openpyxl

#!/usr/bin/python
# -*- coding: cp65001 -*-

r = requests.get("https://www.nsrrc.org.tw/chinese/organizationItem2.aspx?Dept_UID=6",verify=False)
soup = BeautifulSoup(r.text,'html.parser') #將網頁資料以
enterprise_tag = soup.title #找公司名
enterprise = enterprise_tag.string
enterprise = enterprise.replace("\t","")
enterprise = enterprise.replace("\r\n","")
print(enterprise)
groupurl_list=[]
#data = { '全名': [], '名字':[],'中間名':[],'姓氏':[],'電子郵件':[email_item],'公司名稱':[enterprise],'商務電話':[call_item],'學院':[],'系所':[],'行政單位':[groupname]}
titles = ('全名','名字','中間名','姓氏','電子郵件','公司名稱','商務電話','學院','系所','行政單位一級')
wb = openpyxl.Workbook()
sheet = wb.create_sheet(enterprise, 0)
sheet.append(titles)

for i in range(7) :
    #count=0
    id_group = 'ContentPlaceHolder1_dlShowData3_Hyperlink3_'+str(i) #ID找第二層組織(小組)
    all_a = soup.find(id=id_group) 
    groupname = all_a.string #取得小組Name

    #取得小組成員
    groupurl=all_a.get('href')#取得小組頁面
    print(groupname,groupurl)

    url = 'https://www.nsrrc.org.tw/chinese/'
    t_url = url+groupurl
    r2 = requests.get(t_url) #進入小組頁面
    r2.encoding='utf-8'
    group_page = BeautifulSoup(r2.text,'html.parser') #取得各小組頁面 #可使用find、get等功能
    
    driver = webdriver.Chrome()
    
    for _ in range(15):
        per_id = 'ContentPlaceHolder1_dgShowData_hlkUrl_'+str(_)
        if  group_page.find(id=per_id) : #小組組員欄位資料
            per_url = group_page.find(id=per_id).get("href")          
            print(per_url)
            p_url = url+str(per_url)
            print(p_url)
             
            per_page = driver.get(p_url)#進入個人頁面
            name_item = driver.find_element(By.ID,'ContentPlaceHolder1_labEmplName').text#name
            call_item = driver.find_element(By.ID,'ContentPlaceHolder1_labExt').text#分機
            if call_item !="":
                call_item = "03-578-0281#"+call_item
            email_item = driver.find_element(By.ID,'ContentPlaceHolder1_hlkMail').text#email
            print(name_item,call_item,email_item)
            #name = list(name_item)
            #print(name)
            
            ap_list=[]
            ###list index out of range
            name_item=name_item.replace(")","")
            name = name_item.split() #空格切格姓名 0中文全名 1中譯英文名字 2中譯英文姓氏
            name_m=name[1].split("-") #切割單名 
            c_name = name[0].replace("(","") #中文姓氏
            if (len(name_m)>=2): 
                ap_list = [c_name,name_m[0],name_m[1],name[2] ,email_item, enterprise, call_item,"","",groupname] 
            if (len(name)>3 and len(name_m)<2 and len(name_m)>0 ):
                ap_list = [c_name,name[1],name[2],name[3] ,email_item, enterprise, call_item,"","",groupname]
            elif (len(name_m)<2 and len(name_m)>0):
                ap_list = [c_name,name_m[0],"",name[2] ,email_item, enterprise, call_item,"","",groupname]
            
            
                #df=df.append({'全名': name_item, '名字':name[3],'中間名':name[2],'姓氏':name[1],'電子郵件':email_item,'公司名稱':enterprise,'商務電話':call_item,'學院':'','系所':'','行政單位':groupname} , ignore_index=True)
            filename='聯絡人_'+enterprise+'.xlsx'
            sheet.append(ap_list)
            wb.save(filename)
            time.sleep(3)

    
    
    

