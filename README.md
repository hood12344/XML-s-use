# XML-s-use
XML試作
#!/usr/bin/env python
# -*- coding=utf-8 -*-
from xml.etree import ElementTree
import sys
import xml.etree.ElementTree as ET
import urllib.request as httplib
#  SSL  處理，  https    SSSSSS 就需要加上以下2行
import ssl
ssl._create_default_https_context = ssl._create_unverified_context    # 因.urlopen發生問題，將ssl憑證排除

######### 由網路下載 XML 的 字串
# https://data.gov.tw/dataset/67534
# 新竹市里長名冊111.04.07更新
url="https://odws.hccg.gov.tw/001/Upload/25/opendata/9059/135/7b08bec6-6013-48e9-8f4c-8cff183ee35d.xml"
req = httplib.Request( url,data=None,
    headers={ 'User-Agent':"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.88 Safari/537.36"})
reponse = httplib.urlopen(req)               # 開啟連線動作
if reponse.code==200:                        # 當連線正常時
    contents=reponse.read()                  # 讀取網頁內容
    contents=contents.decode("utf-8")        # 轉換編碼為 utf-8
    print(contents)
    # 把取得的XML 資料，存在檔案中
    fr = open('新竹市里長名冊111.04.07更新.xml', 'w', encoding="utf-8")
    fr.write(contents)
    fr.close()



# 加载XML文件
root = ElementTree.fromstring(contents)
data= root.findall("Data")


# 取得Tags
elemList = []
for elem in root.iter():
    elemList.append(elem.tag)

elemList = list(set(elemList))
print(elemList)


# 取得所有資料
yearList=[]
titleList=[]
n=0
while n<len(data):
    # XML 解析
    year           = data[n].findall("屆別")
    name           = data[n].findall("姓名")
    title          = data[n].findall("職稱")
    gender          = data[n].findall("性別")
    Area            = data[n].findall("區別")
    Ribe            =data[n].findall("里別")
    mailing_address =data[n].findall("通訊地址")
    Local_calls     =data[n].findall("市內電話")
    cell_phone      =data[n].findall("行動電話")
    str1=" 屆別:"+year[0].text+\
        " ,姓名："+name[0].text+\
        " ,職稱："+title[0].text+\
        " ,性別："+gender[0].text+\
        " ,區別："+Area[0].text+\
        " ,里別:"+Ribe[0].text+\
        " ,通訊地址"+mailing_address[0].text+\
        " ,市內電話"+Local_calls[0].text+\
        " ,行動電話"+cell_phone[0].text


    yearList.append(year[0].text)
    titleList.append(title[0].text)
    print(str1)
    n=n+1





import matplotlib.pyplot as plt # 匯入matplotlib 的pyplot 類別，並設定為plt
from matplotlib.font_manager import FontProperties # 中文字體

# 換成中文的字體
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
plt.rcParams['axes.unicode_minus'] = False  # 步驟二（解決座標軸負數的負號顯示問題）



plt.plot(yearList,titleList, "g--",label='上市公司-資本額')
plt.legend()         # 自動改變顯示的位置

plt.title("國內公開發行公司股票年度發行概況")
plt.ylabel('上市公司-資本額') # 顯示Y 座標的文字
plt.xlabel('年度') # 顯示Y 座標的文字

plt.show() # 繪製

