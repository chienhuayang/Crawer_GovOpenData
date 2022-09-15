# <font size=4>行政院農業委員會種苗改良繁殖場</font></br>  <font size=3>總共50筆人員資料,其中一筆姓名無值忽略</font></br><font size=3>包含：繁殖技術課、品種改良保護課、種苗經營課、生物技術課、技術服務室、農場、本場、行政室、主計機構、人事機構</font></br></br> <font size=3>無email資料</font></br>

<font size=3>Raw data為每個值一行 使用if else抓關鍵字作為資料分類依據</font></br></br>

 <font size=3>找出一級行政單位名使用driver.find_element(By.XPATH,'//div//h2'):
 groupname=driver.find_element(By.XPATH,'//div//h2').text
 </font></br></br>
 <font size=3>使用driver.find_elements(By.XPATH,"//tbody//tr//td")</font></br> <font size=4>